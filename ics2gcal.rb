#!/usr/bin/ruby
# -*- coding: utf-8 -*-
#
# Takes a .ics file as parameter, adds it to Google Calendar
#
# (C) 2014 INOUE Tomohiro
#
# Licensed under MIT License, see included file LICENSE

### 個人毎にカスタマイズが必要
# 同期先 Google Calendar のカレンダー名
CalendarName = "仕事用" 
# ics ファイルで除外するカテゴリのリスト
ExcludeCategories = [ "祝日" ] 

require 'date'
require 'json'
require 'logger'

require 'rubygems'
require 'bundler/setup'

require 'icalendar'
require 'icalendar/tzinfo'
require 'base32'
require 'google/api_client'
require 'google/api_client/client_secrets'
require 'google/api_client/auth/file_storage'
require 'google/api_client/auth/installed_app'
require 'logger'

logger = Logger.new(STDOUT)
logger.level = Logger::INFO

logger.debug("Created logger")
logger.info("Program started")

logger.debug ENV["http_proxy"] 
logger.debug ENV["https_proxy"] 
logger.debug ENV["LANG"] 

### util functions
def datetime_hash(datetime, timezone)
  hash = {}
  #p datetime.class
  if datetime.instance_of?(Icalendar::Values::DateTime)
    return {"dateTime" => datetime.iso8601, "timeZone" => timezone}
  elsif datetime.instance_of?(Icalendar::Values::Date)
    return {"date" => datetime.iso8601, "timeZone" => timezone}
  else 
    return nil
  end
end 

if ARGV.length < 1
  abort "Need at least 1 .ics file"
  exit
end

### Read ics file
ics_file = File.open(ARGV[0])
icalendars = Icalendar::parse(ics_file)

TimeZone = 'Asia/Tokyo'

# customize Base32 for compat to GCalendar (RFC2938)
Base32.table = 'abcdefghijklmnopqrstuv0123456789'.freeze

# convert ics events to GCalendar events
@events = {}
icalendars.first.events.each do |event|
  logger.debug(event)
  gcevent_id = Base32.encode(event.uid).gsub(%r|=+|,'').downcase

  gcevent = {}
  gcevent["id"]       = gcevent_id
  gcevent["summary"]  = event.summary
  gcevent["start"]    = datetime_hash(event.dtstart, TimeZone)
  gcevent["end"]      = datetime_hash(event.dtend, TimeZone)
  gcevent["categories"] = event.categories.to_a

  @events[gcevent_id] = gcevent
end


### Setup Google API

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
  logger.debug c["id"]
  if c["summary"] == CalendarName
    @cal_id = c["id"]
    break
  end
end

abort("cant find calendar") if @cal_id.nil?
logger.debug @cal_id

# Google Calendar 登録済みイベントのスキャン
day_from = DateTime.now - 365 # 一年前から
day_to   = DateTime.now + 365 # 一年後まで

params = {}
params['calendarId'] = @cal_id
params['timeMin'] = day_from.iso8601
params['timeMax'] = day_to.iso8601
logger.debug params

# Page loop...
gcevents = client.execute(:api_method => service.events.list,:parameters => params)
while true
  gcevents.data.items.each do |e|
    gcevent_id = e['id']
    gcevent_summary = e["summary"]

    # ics にあるイベントは更新、無いイベントは削除
    gcevent = @events[gcevent_id]
    if gcevent then
      logger.info "Updating #{gcevent_summary}..."
      result2 = client.execute(:api_method => service.events.patch,
                               :parameters => {'calendarId' => @cal_id, 
                                 'eventId' => gcevent_id},
                               :body => JSON.dump(gcevent),
                               :headers => {'Content-Type' => 'application/json'})
      if JSON.parse(result2.response.body).has_key?("error") then
        logger.info result2.response.body
      end
    else # ics に無いイベントは削除
      logger.info "Deleting #{gcevent_summary}..."
      client.execute(:api_method => service.events.delete,
                     :parameters => {'calendarId' => @cal_id, 'eventId' => gcevent_id})
    end
    #処理済みイベントをメモリから削除（後で追加しないように）
    @events.delete(gcevent_id)
  end
  if !(page_token = gcevents.data.next_page_token)
    break
  end
  params["pageToken"] = page_token
  gcevents = client.execute(:api_method => service.events.list,:parameters => params)
end

# Google Calendar に無かったイベントを追加
@events.each do |gcevent_id, gcevent|
  logger.debug gcevent_id
  logger.debug gcevent

  date = gcevent["start"]["dateTime"] || gcevent["start"]["date"]
  logger.debug "\n"
  logger.info "Adding event ... "
  logger.debug gcevent_id
  logger.info "#{gcevent["summary"]} on #{date}"
  
  if gcevent["categories"] & ExcludeCategories != []  # 積集合が空じゃない
    logger.info "skipped (reason: excluded category)"
    next
  end

  result = client.execute(:api_method => service.events.insert,
                          :parameters => {'calendarId' => @cal_id},
                          :body => JSON.dump(gcevent),
                          :headers => {'Content-Type' => 'application/json'})

  result_hash = JSON.parse(result.response.body)
  if result_hash.has_key?("error") then
    #logger.debug "event duplicated" if (result_hash["error"]["errors"].first["reason"] == "duplicate") 
    logger.warn("Event add failed.")
    logger.warn(result.response.body)
    logger.debug result.request
  end
end

logger.info "Program end."
