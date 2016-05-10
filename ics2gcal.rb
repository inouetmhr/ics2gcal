#!/usr/local/bin/ruby
# -*- coding: utf-8 -*-
#
# One way sync tool from iCalendar (ics file) to Google Calendar
# Takes a .ics file as parameter, adds it to Google Calendar
# 
# (C) 2014 INOUE Tomohiro < tm.inoue (gmail)>
#
# Licensed under MIT License, see included file LICENSE
#
# ---
# variables difinitioan:
#  - ievent : an instance of iCalendar event (including recurring event)
#  - ievent_ex : an exception (a modified instance) of recurring evnent
#  - gevent : an instance of Google calendar event (including recurring event)
#  - gitem  : each event item of Google calendar's recurring event
# +++

require 'date'
require 'json'
require 'logger'
require 'optparse'

require 'rubygems'
require 'bundler/setup'

require 'tzinfo'

require 'icalendar'
require 'icalendar/tzinfo'
require 'base32'

require 'googleauth'
require 'googleauth/stores/file_token_store'
require 'google/apis/calendar_v3'
require 'google/api_client/client_secrets'

GEvent = Google::Apis::CalendarV3::Event

@logger = Logger.new(STDOUT)
@logger.level = Logger::INFO
#@logger.level = Logger::DEBUG
@logger.level = Logger::DEBUG if $DEBUG
@logger.debug("Created logger")

# 
# Default dispalay name of Google Calendar to be synchronized
@calendarName ||= "仕事用" 
# List of categories that are excluded 
@excludeCategories ||= [ "祝日" ] 
opt = OptionParser.new
opt.on('-g CALENDAR', 'Name of Google Calendar to be synchronized') {|v| 
  @calendarName = v 
}
opt.on('-x CATEGORIES', 'Comma seperated list of categories that will not be sync'){|v| 
  @excludeCategories = v.split(',')
}
opt.on('-V', 'puts verbose log messages'){|v| @logger.level = Logger::DEBUG}

opt.parse!(ARGV)
#p @calendarName, @excludeCategories, ARGV

if ARGV.length < 1
  abort "Need one .ics file in argments"
else
  ICSFILE = ARGV[0]
end

@logger.info("Program started")
@logger.debug ENV["http_proxy"] 
@logger.debug ENV["https_proxy"] 
@logger.debug ENV["LANG"] 

#default time zone
TimeZone = 'Asia/Tokyo'
# customize Base32 for compatible to Google Calendar (RFC2938)
Base32.table = 'abcdefghijklmnopqrstuv0123456789'.freeze


### util functions

# @return [Google::Apis::CalendarV3::EventDateTime]
def get_date(event)
  if event.kind_of?(GEvent) then
    return event.start || event.original_start_time
  else
    return event[:start] || event[:original_start_time]
  end  
end

# @return String 
def get_date_string(event)
  dt = get_date(event)
  return (dt.date_time || dt.date).to_s
end

# To fix timezone:
# iCalendar gem library ignores tzid (time zone) on parsing ics_file
class Icalendar::Values::DateTime
  def fix_JST  # minus 9 hours 
    return (self - Rational(9,24)).new_offset("+0900")
  end

  def fix_offset(offset)  # offset: e.g. Rational(9,24) 
    return (self - offset).new_offset(offset)
  end
end

class Icalendar::Event
  def fix_timezone!
    tzids = @dtstart.ical_params["tzid"]
    return unless tzids # UTC の場合を想定 (未検証!)

    #TODO tzids が複数の場合未検証
    tzid = tzids.first
    offset = @parent.timezone_offset(tzid)
    if offset then
      @dtstart = @dtstart.fix_offset(offset)
      @dtend   = @dtend.fix_offset(offset)
    else
      Icalendar.logger.warn "TZID #{tzid} not processed" 
    end
    return self
  end
end

class Icalendar::Calendar
  def utc_offset_to_rational(utc_offset)
    offset_seconds = utc_offset.hours * 60 * 60 + \
                     utc_offset.minutes * 60 + \
                     utc_offset.seconds
    offset_seconds = - offset_seconds if utc_offset.behind 
    Rational(offset_seconds, 24 * 60 * 60 )
  rescue
    nil
  end

  # return timezone offset: e.g. JST-9 => Rational(9/24)
  def timezone_offset(tzid)
    @timezones.each{|tz|
      next unless tz.tzid == tzid

      #FIXME standards.first.tzoffsetto が妥当なのか不明
      # daylight (夏時間?) ありだとだめかも
      return utc_offset_to_rational(tz.standards.first.tzoffsetto)
    }
    return nil
  end
end

# convert Date/DateTime object to a Hash object that represents date/time
def datetime_hash(datetime, timezone)
  case datetime
  when DateTime, Icalendar::Values::DateTime
    return {"dateTime" => datetime.iso8601, "timeZone" => timezone}
  when Date, Icalendar::Values::Date
    return {"date" => datetime.iso8601, "timeZone" => timezone}
  else 
    raise "datetime_hash class error"
  end
end 

# @return Google::Apis::CalendarV3::EventDateTime
def event_datetime(datetime, timezone)
  hash = {}
  #p datetime.class
  case datetime
  when DateTime, Icalendar::Values::DateTime
    hash = {:date_time => datetime.iso8601, :time_zone => timezone}
  when Date, Icalendar::Values::Date
    hash = {:date => datetime.iso8601, :time_zone => timezone}
  else 
    raise "event_datetime class error"
  end
  return Google::Apis::CalendarV3::EventDateTime.new(**hash)
end 

def delete_event(event_id)
  begin
    @client.delete_event(@cal_id, event_id)
  rescue Google::Apis::ClientError => e
    @logger.warn "event delete failed."
    @logger.warn e.messages
    @logger.debug e
  end
end

### Read iCalendar (ics file)
@logger.info("Loading ics file #{ICSFILE} ...")
icalendars = nil
File.open(ICSFILE){|f| icalendars = Icalendar::parse(f) }

# convert ics events to GCalendar events
@events = {}
@recurrence_exceptions = {}
icalendars.first.events.each do |ievent|
  @logger.debug(ievent)
  
  gevent_id = Base32.encode(ievent.uid).gsub(%r|=+|,'').downcase
  
  gevent = {}
  gevent[:id]      = gevent_id
  gevent[:summary] = ievent.summary
  gevent[:categories] = ievent.categories.to_a

  
  #繰り返し予定
  if ievent.rrule != [] then
    recurrence = []
    ievent.rrule.each{|e| recurrence << "RRULE:" + e.value_ical }
    #recurrence << ievent.exrule if ievent.exrule
    if ievent.exdate != [] then
      ievent.exdate.each{|e| recurrence << "EXDATE:" + e.value_ical }
    end
    if ievent.rdate != [] then
      ievent.rdate.each{|e| recurrence << "RDATE:" + e.value_ical }
    end
    @logger.debug "Recurrence - " + recurrence.to_s
    gevent[:recurrence] = recurrence

  end

  #繰り返しの例外（変更）の対処
  #UIDが同じなので別の処理が必要
  #Google Calendar のインタフェースも異なる
  if ievent.recurrence_id then
    gevent[:recurrence_id] = ievent.recurrence_id
    gevent[:summary] ||= @events[gevent_id][:summary]
    
#    gevent[:start] = datetime_hash(ievent.dtstart, TimeZone)
#    gevent[:end]   = datetime_hash(ievent.dtend, TimeZone)
    gevent[:start]   = event_datetime(ievent.dtstart, TimeZone)
    gevent[:end]     = event_datetime(ievent.dtend, TimeZone)
    #gevent["start"]   ||= @events[gevent_id]["start"]
    #gevent["end"]     ||= @events[gevent_id]["end"]
    
    @logger.debug ievent.summary
    @logger.debug ievent.dtstart
    @logger.debug gevent[:start]

    @recurrence_exceptions[gevent_id] ||= [] # nil の場合
    @recurrence_exceptions[gevent_id] << gevent
    next
  end

  # icsファイルの tzid を正しく読まないので、終日イベント以外は修正
  # 繰り返しのイベントはなぜ修正しなくてよいの???
  if ievent.dtstart.instance_of?(Icalendar::Values::DateTime)
    ievent.fix_timezone!
  end
  # TODO Google に TimeZone 渡さないといけないんだっけ? （そうだった気がするけど）
#  gevent[:start]   = datetime_hash(ievent.dtstart, TimeZone)
#  gevent[:end]     = datetime_hash(ievent.dtend, TimeZone)
  gevent[:start]   = event_datetime(ievent.dtstart, TimeZone)
  gevent[:end]     = event_datetime(ievent.dtend, TimeZone)

  @events[gevent_id] = gevent

  @logger.debug ievent.summary
  @logger.debug ievent.dtstart
  @logger.debug gevent[:start]
end

#exit   
### Setup Google API
#TODO bulk update

# Initialize the client.
# OAuth2.0 auth and  cache 
OOB_URI = 'urn:ietf:wg:oauth:2.0:oob'
CREDENTIAL_STORE_FILE = "#{$0}-tokens.yaml"
scope = 'https://www.googleapis.com/auth/calendar'

client_id = Google::Auth::ClientId.from_file('./client_secrets.json')
token_store = Google::Auth::Stores::FileTokenStore.new(
  :file => CREDENTIAL_STORE_FILE)
authorizer = Google::Auth::UserAuthorizer.new(client_id, scope, token_store)

user_id = 'default'
credentials = authorizer.get_credentials(user_id)
if credentials.nil?
  url = authorizer.get_authorization_url(base_url: OOB_URI )
  puts "Open #{url} in your browser and enter the resulting code:"
  code = STDIN.gets
  credentials = authorizer.get_and_store_credentials_from_code(
    user_id: user_id, code: code, base_url: OOB_URI)
end  

# Initialize Google API. Note this will make a request to the
# discovery service every time, so be sure to use serialization
# in your production code. Check the samples for more details.
@client = Google::Apis::CalendarV3::CalendarService.new
@client.authorization = credentials

#カレンダーリストの取得
gcalendars = @client.list_calendar_lists

@cal_id = nil
gcalendars.items.each do |c|
  @logger.debug c.summary
  if c.summary == @calendarName
    @cal_id = c.id
    break
  end
end

abort("Could not find Google Calendar: #{@calendarName}") if @cal_id.nil?
@logger.debug @cal_id

# Google Calendar 登録済みイベントのスキャン
day_from = DateTime.now - 365 # 一年前から
day_to   = DateTime.now + 365 # 一年後まで

def gc_update(gevent, ievent)
  #GCカレンダーアイテムの更新
  gevent_id = gevent.id
  gevent_summary = gevent.summary
  @logger.info "Updating #{gevent_summary} on #{get_date_string(gevent)}"
  
  ## TODO ievent から CalendarV3::Event に変換できてる？
  @logger.debug ievent
  gevent2 = @client.patch_event(@cal_id, gevent_id, GEvent.new(**ievent))
  unless gevent2
    @logger.warn "gc_update failed?"
  end
  #                                :body => JSON.dump(ievent),
  #                                :headers => {'Content-Type' => 'application/json'})
  #  if JSON.parse(result2.response.body).has_key?("error") then
  #    @logger.info result2.response.body
  #    @logger.warn result2.request.body
  #  end
end

# GCイベントをスキャン with Pagerization
gevents = @client.list_events(@cal_id, \
                              time_min: day_from.iso8601, time_max: day_to.iso8601)
while true
#  gevents.data.items.each do |gevent|
  gevents.items.each do |gevent|
    #gevent_id = gevent['iCalUID']
    #gevent_id = gevent_id ? gevent_id.gsub(/@google.com$/, "")  : gevent['id']
    gevent_id = gevent.id
    gevent_summary = gevent.summary

    ievent = @events[gevent_id]
    if ievent then
      # gc の id が ics のid と一致したイベントは、更新する
      # 現状、全部更新してる
      gc_update(gevent, ievent)
    else # gc と ics で id が一致しない（icsに無い）イベントはケースに応じて処理
      @logger.debug gevent
      if gevent.recurring_event_id then ## 空のとき nil ？ TODO
        # 繰り返しイベントのインスタンスの場合はスキップ
        @logger.info "Skipping Recurring event on #{get_date_string(gevent)} ..."
      else 
        # そうでない場合は、icsから削除されたイベントか、別IDで登録さ
        # れた例外イベントなので、削除する（変更は追加で対応）
        @logger.info "Deleting #{gevent_summary} on #{get_date_string(gevent)} ..."
        delete_event(gevent_id)
      end
    end
    #処理済みイベントをメモリから削除（後で追加しないように）
    @events.delete(gevent_id)
  end
  if !(page_token = gevents.next_page_token)
    break
  end
  gevents = @client.list_events(@cal_id, \
                                time_min: day_from.iso8601, time_max: day_to.iso8601, \
                                page_token: page_token)
end

# Google Calendar に無かったイベントを追加
@events.each do |gevent_id, gevent|
  @logger.debug gevent_id
  @logger.debug gevent

  @logger.debug "\n"
  @logger.info "Adding event ... "
  @logger.debug gevent_id
  @logger.info "#{gevent[:summary]} on #{get_date_string(gevent)}"
  
  if gevent[:categories] & @excludeCategories != []  # 積集合が空じゃない
    @logger.info "skipped (reason: excluded category)"
    next
  end

  begin
    @client.insert_event(@cal_id, GEvent.new(gevent))
  #if result_hash.has_key?("error") then
  rescue Google::Apis::ClientError => e
    @logger.debug e.status_code
    @logger.debug e.body
    #if (result_hash["error"]["errors"].first["reason"] == "duplicate") then
    result = JSON.parse(e.body)
    if result["error"]["errors"][0]["reason"] == "duplicate" then
      # なんらかの理由で、イベントは削除されたが gc の id が残っている状態の場合
      @logger.warn "event duplicated (id = #{gevent_id})"
      @logger.debug gevent
      # gc_update(gevent, gevent) #これだと更新は成功するが削除状態のままになる

      # しかたないので id を空にして新規追加する
      # TODO: これだと次回実行時にまた削除されて毎回追加になる
      gevent.delete(:id)
      @logger.warn "Retry adding with null id"
      redo
    else
      @logger.warn("Event add failed.")
      @logger.warn e.message
      @logger.warn e.status_code
      @logger.warn e.body
    end
  end
end

## 繰り返しイベントの例外（変更）対応
# ics例外イベントのスキャン
@recurrence_exceptions.each do |gevent_id, exceptions|

  @logger.debug gevent_id
  @logger.debug exceptions

  original_dates = exceptions.map{|i| i[:recurrence_id]}
  @logger.debug original_dates

  #exceptions.each{|e| p gevent["start"] ; p gevent["recurrence_id"]}
  # ソート済みの前提
  
  #gc 個別イベント (gitem) のスキャン
  #これには例外だけで無く全てのイベントインスタンスが含まれる
  gc_items = @client.list_event_instances(@cal_id, gevent_id)
  # FIXME handle error case
  #gc_items = JSON.parse(result.body)["items"]
  gc_items.items.each{|gitem|
    # （方針） gitem の更新がなぜか上手くいかないため、gitemはGCから削
    # 除して、別のUIDのイベントとして登録する。

    #ToDo DateTime じゃなく Date のケースのテスト
    orig_start = gitem.original_start_time
    #orig_dt =  DateTime.iso8601(orig_start["dateTime"] || orig_start["date"])
    orig_dt =  orig_start
    @logger.debug orig_dt

    # ics にも変更元がある gc 例外イベントの場合
    if original_dates.index(orig_dt) then
      # 更新したいが上手くいかない
      # gevent = exceptions[original_dates.index(orig_dt)]
      # gevent.delete('id')
      # @logger.info JSON.dump(gevent)
      # result2 = @client.execute(:api_method => @service.events.patch
      #                          :parameters => {'calendarId' => @cal_id,
      #                            'eventId' => gitem['id']},
      #                          :body_object => JSON.dump(gevent),
      #                         :headers => {'Content-Type' => 'application/json'})

      # 例外イベントの更新のためにまずGCから削除（あとで追加）
      @logger.info "Deleting instance ... "
      @logger.info "#{gitem[:summary]} on #{get_date_string(gitem)}"
      delete_event(gitem[:id])

   else # ics に変更元がない gc 例外イベントの場合
     # keep it (do nothing)
   end
  }

  # icsの例外イベントをGCに新規イベントとして登録
  exceptions.each{|ievent_ex|
    ievent_ex.delete(:id) 
    # TODO 現状 gitem は削除追加で更新しているが、IDの変換ルールを導入
    # すれば更新ですむかもしれない

    @logger.info "Adding instance ... "
    @logger.info "#{ievent_ex[:summary]} on #{get_date_string(ievent_ex)}"
    begin
      @client.insert_event(@cal_id, GEvent.new(ievent_ex))
    rescue Google::Apis::ClientError => e
      @logger.warn("Instance add failed.")
      @logger.warn(e.body)
      @logger.debug e
    end
  }
end

@logger.info "Program end."
