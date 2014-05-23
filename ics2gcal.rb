#!/usr/bin/ruby
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
#  - gitem  : each event item of a Google calendar recurring event
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
require 'google/api_client'
require 'google/api_client/client_secrets'
require 'google/api_client/auth/file_storage'
require 'google/api_client/auth/installed_app'

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

# get date/time string from a hash representing an Google event/item
def get_date(gevent)
  date = gevent["start"] || gevent["originalStartTime"]
  return date["dateTime"] || date["date"]
end

# iCalendar library (GEM) seems not read tzid (time zone) in ics file
# This hack returns a modified datetime as JST time zone was set
class Icalendar::Values::DateTime
  def fix_JST  # minus 9 hours 
    return (self - Rational(9,24)).new_offset("+0900")
  end
end

# def fix_JST(datetime)
#   return nil if datetime.kind_of?(DateTime)
#   case datetime.zone
#   when "+00:00"
#     return (datetime - Rational(9,24)).new_offset("+0900")
#   else
#     return datetime
#   end
# end

# convert Date/DateTime object to a Hash object that represents date/time
def datetime_hash(datetime, timezone)
  hash = {}
  #p datetime.class
  case datetime
  when DateTime, Icalendar::Values::DateTime
    return {"dateTime" => datetime.iso8601, "timeZone" => timezone}
  when Date, Icalendar::Values::Date
    return {"date" => datetime.iso8601, "timeZone" => timezone}
  else 
    raise "datetime_hash class error"
  end
end 

### Read iCalendar (ics file)
@logger.info("Loading ics file #{ARGV[0]} ...")
icalendars = nil
File.open(ARGV[0]){|f| icalendars = Icalendar::parse(f) }

# convert ics events to GCalendar events
@events = {}
@recurrence_exceptions = {}
icalendars.first.events.each do |ievent|
  @logger.debug(ievent)
  
  gevent_id = Base32.encode(ievent.uid).gsub(%r|=+|,'').downcase
  
  gevent = {}
  gevent["id"]      = gevent_id
  gevent["summary"] = ievent.summary
  gevent["categories"] = ievent.categories.to_a

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
    gevent["recurrence"] = recurrence

  end

  #繰り返しの例外（変更）の対処
  #UIDが同じなので別の処理が必要
  #Google Calendar のインタフェースも異なる
  if ievent.recurrence_id then
    gevent["recurrence_id"] = ievent.recurrence_id
    gevent["summary"] ||= @events[gevent_id]["summary"]
    
    gevent["start"] = datetime_hash(ievent.dtstart, TimeZone)
    gevent["end"]   = datetime_hash(ievent.dtend, TimeZone)
    #gevent["start"]   ||= @events[gevent_id]["start"]
    #gevent["end"]     ||= @events[gevent_id]["end"]
    
    @logger.debug ievent.summary
    @logger.debug ievent.dtstart
    @logger.debug gevent["start"]

    @recurrence_exceptions[gevent_id] ||= [] # nil の場合
    @recurrence_exceptions[gevent_id] << gevent
    next
  end

  # 繰り返しで無い場合、iCalendar の tzid を正しく読まないので修正
  if ievent.dtstart.instance_of?(Icalendar::Values::DateTime)
    gevent["start"]   = datetime_hash(ievent.dtstart.fix_JST, TimeZone)
    gevent["end"]     = datetime_hash(ievent.dtend.fix_JST, TimeZone)
  else 
    gevent["start"]   = datetime_hash(ievent.dtstart, TimeZone)
    gevent["end"]     = datetime_hash(ievent.dtend, TimeZone)
  end
  @events[gevent_id] = gevent

  @logger.debug ievent.summary
  @logger.debug ievent.dtstart
  @logger.debug gevent["start"]
end

#exit   
### Setup Google API
#TODO bulk update

# Initialize the client.
client = Google::APIClient.new(
                               :application_name => 'ICS to Google Calendar',
                               :application_version => '0.1.0'
                               )

# OAuth2.0 auth and  cache 
CREDENTIAL_STORE_FILE = "#{$0}-oauth2.json"
file_storage = Google::APIClient::FileStorage.new(CREDENTIAL_STORE_FILE)

if file_storage.authorization.nil?
  # Load client secrets from your client_secrets.json.
  client_secrets = Google::APIClient::ClientSecrets.load
  flow = Google::APIClient::InstalledAppFlow.new(
                                                 :client_id => client_secrets.client_id,
                                                 :client_secret => client_secrets.client_secret,
                                                 :scope => ['https://www.googleapis.com/auth/calendar']
                                                 )
  
  client.authorization = flow.authorize(file_storage)
else
  client.authorization = file_storage.authorization
end


# Initialize Google API. Note this will make a request to the
# discovery service every time, so be sure to use serialization
# in your production code. Check the samples for more details.
service = client.discovered_api('calendar', 'v3')

#カレンダーリストの取得
gcalendars = client.execute(:api_method => service.calendar_list.list)

@cal_id = nil
gcalendars.data.items.each do |c|
  @logger.debug c["id"]
  if c["summary"] == @calendarName
    @cal_id = c["id"]
    break
  end
end

abort("Could not find Google Calendar: #{@calendarName}") if @cal_id.nil?
@logger.debug @cal_id

# Google Calendar 登録済みイベントのスキャン
day_from = DateTime.now - 365 # 一年前から
day_to   = DateTime.now + 365 # 一年後まで

params = {}
params['calendarId'] = @cal_id
params['timeMin'] = day_from.iso8601
params['timeMax'] = day_to.iso8601
@logger.debug params

# GCイベントをスキャン with Pagerization
gevents = client.execute(:api_method => service.events.list,:parameters => params)
while true
  gevents.data.items.each do |gevent|
    #gevent_id = gevent['iCalUID']
    #gevent_id = gevent_id ? gevent_id.gsub(/@google.com$/, "")  : gevent['id']
    gevent_id = gevent['id']
    gevent_summary = gevent["summary"]

    ievent = @events[gevent_id]
    # gc の id が ics のid と一致したイベントは、更新する
    if ievent then
      @logger.info "Updating #{gevent_summary} on #{get_date(gevent)}"
      result2 = client.execute(:api_method => service.events.patch,
                               :parameters => {'calendarId' => @cal_id, 
                                 'eventId' => gevent_id},
                               :body => JSON.dump(ievent),
                               :headers => {'Content-Type' => 'application/json'})
      if JSON.parse(result2.response.body).has_key?("error") then
        @logger.info result2.response.body
      end
    else # gc と ics で id が一致しないイベントは、ケースに応じて処理
      if gevent['recurringEventId'] then
        # 繰り返しイベントのインスタンスの場合はスキップ
        @logger.info "Skipping Recurring event on #{get_date(gevent)} ..."
      else 
        # そうでない場合は、icsから削除されたイベントか、別IDで登録さ
        # れた例外イベントなので、削除する（変更は追加で対応）
        @logger.info "Deleting #{gevent_summary} on #{get_date(gevent)} ..."
        client.execute(:api_method => service.events.delete,
                       :parameters => {'calendarId' => @cal_id, 'eventId' => gevent_id})
      end
    end
    #処理済みイベントをメモリから削除（後で追加しないように）
    @events.delete(gevent_id)
  end
  if !(page_token = gevents.data.next_page_token)
    break
  end
  params["pageToken"] = page_token
  gevents = client.execute(:api_method => service.events.list,:parameters => params)
end

# Google Calendar に無かったイベントを追加
@events.each do |gevent_id, gevent|
  @logger.debug gevent_id
  @logger.debug gevent

  @logger.debug "\n"
  @logger.info "Adding event ... "
  @logger.debug gevent_id
  @logger.info "#{gevent["summary"]} on #{get_date(gevent)}"
  
  if gevent["categories"] & @excludeCategories != []  # 積集合が空じゃない
    @logger.info "skipped (reason: excluded category)"
    next
  end

  result = client.execute(:api_method => service.events.insert,
                          :parameters => {'calendarId' => @cal_id},
                          :body => JSON.dump(gevent),
                          :headers => {'Content-Type' => 'application/json'})

  result_hash = JSON.parse(result.response.body)
  if result_hash.has_key?("error") then
    #@logger.debug "event duplicated" if (result_hash["error"]["errors"].first["reason"] == "duplicate") 
    @logger.warn("Event add failed.")
    @logger.warn(result.response.body)
    @logger.debug result.request
  end
end

## 繰り返しイベントの例外（変更）対応
# ics例外イベントのスキャン
@recurrence_exceptions.each do |gevent_id, exceptions|

  @logger.debug gevent_id
  @logger.debug exceptions

  original_dates = exceptions.map{|i| i['recurrence_id']}
  @logger.debug original_dates

  #exceptions.each{|e| p gevent["start"] ; p gevent["recurrence_id"]}
  # ソート済みの前提
  
  #gc 個別イベント (gitem) のスキャン
  #これには例外だけで無く全てのイベントインスタンスが含まれる
  result = client.execute(:api_method => service.events.instances,
                             :parameters => {'calendarId' => @cal_id,
                               'eventId' => gevent_id})
  # FIXME handle error case
  gc_items = JSON.parse(result.body)["items"]
  gc_items.each{|gitem|
    # （方針） gitem の更新がなぜか上手くいかないため、gitemはGCから削
    # 除して、別のUIDのイベントとして登録する。

    #ToDo DateTime じゃなく Date のケースのテスト
    orig_start = gitem["originalStartTime"]
    orig_dt =  DateTime.iso8601(orig_start["dateTime"] || orig_start["date"])
    @logger.debug orig_dt

    # ics にも変更元がある gc 例外イベントの場合
    if original_dates.index(orig_dt) then
      # 更新したいが上手くいかない
      # gevent = exceptions[original_dates.index(orig_dt)]
      # gevent.delete('id')
      # @logger.info JSON.dump(gevent)
      # result2 = client.execute(:api_method => service.events.patch
      #                          :parameters => {'calendarId' => @cal_id,
      #                            'eventId' => gitem['id']},
      #                          :body_object => JSON.dump(gevent),
      #                         :headers => {'Content-Type' => 'application/json'})

      # 例外イベントの更新のためにまずGCから削除（あとで追加）
      @logger.info "Deleting instance ... "
      @logger.info "#{gitem["summary"]} on #{get_date(gitem)}"
      result = client.execute(:api_method => service.events.delete,
                              :parameters => {'calendarId' => @cal_id, 
                                'eventId' => gitem['id']})
      if result.response.body != "" then
        result_error = JSON.parse(result.response.body)["error"] 
        if result_error then
          @logger.warn("Instance delete failed.")
          @logger.warn(result.response.body)
          @logger.debug result.request
        end
      end

   else # ics に変更元がない gc 例外イベントの場合
     # keep it (do nothing)
   end
  }

  # icsの例外イベントをGCに新規イベントとして登録
  exceptions.each{|ievent_ex|
    ievent_ex.delete('id') 
    # TODO 現状 gitem は削除追加で更新しているが、IDの変換ルールを導入
    # すれば更新ですむかもしれない

    @logger.info "Adding instance ... "
    @logger.info "#{ievent_ex["summary"]} on #{get_date(ievent_ex)}"
    result = client.execute(:api_method => service.events.insert,
                            :parameters => {'calendarId' => @cal_id},
                            :body => JSON.dump(ievent_ex),
                            :headers => {'Content-Type' => 'application/json'})
    if JSON.parse(result.response.body).has_key?("error") then
      @logger.warn("Instance add failed.")
      @logger.warn(result.response.body)
      @logger.debug result.request
    end
  }
end

@logger.info "Program end."
