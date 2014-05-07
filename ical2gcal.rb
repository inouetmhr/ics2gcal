#!/usr/bin/ruby
#
# Takes a .ics file as parameter, adds it to Google Calendar
#
# Change USERNAME and PASSWORD to your username and password, respectively,
# set DEFAULT_CAL to define the calendar selected by default (you'll still
# get the option to choose one).
#
# (C) 2009 David Verhasselt (david@crowdway.com)
# (C) 2014 INOUE Tomohiro
#
# Licensed under MIT License, see included file LICENSE

USERNAME = "tm.inoue"
DEFAULT_CAL = "test"

require 'rubygems'
require 'icalendar'
require 'googlecalendar'

# Make "calendars" accessible
class Googlecalendar::GData
	attr_reader :calendars
end

class GoogleCalendarAPI
  #
end

if ARGV.length < 1
	puts "Need at least 1 .ics file"
	exit
end

cal_file = File.open(ARGV[0])
cals = Icalendar::parse(cal_file)

cals.first.events.each do |event|
   puts "summary: " + event.summary
   puts "start: " + event.dtstart.to_s
   puts "end: " + event.dtend.to_s
  puts 
end


g = Googlecalendar::GData.new;
g.login(USERNAME, PASSWORD);
g.get_calendars();
cals = g.calendars.map { |item| item.title }

p cals

exit

eventHash[:title] = event.summary
eventHash[:content] = event.description
eventHash[:where] = event.location
eventHash[:startTime] = event.start_date
eventHash[:endTime] = event.end_date
eventHash[:author] = "iCal2GCal"
eventHash[:email] = "iCal2GCal"

g.new_event(eventHash, calSelector.active_text)


