// Room Reservation System Video
// Kurt Kaiser, 2018
// All rights reserved

var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheets()[0];
var lastRow = sheet.getLastRow();
var lastColumn = sheet.getLastColumn();

// Calendars to output appointments to
var cal101 = CalendarApp.getCalendarById('hl39b5l1bapd3ul6cpkt6s0dfo@group.calendar.google.com');
var cal102 = CalendarApp.getCalendarById('ci4uov4s8eap771v9hpr70jp60@group.calendar.google.com');
var cal103 = CalendarApp.getCalendarById('svbrato5stpi7f5grrd6ha4924@group.calendar.google.com');
var cal202 = CalendarApp.getCalendarById('8iagfnive672quk59ukkaq0rfc@group.calendar.google.com');
var cal103 = CalendarApp.getCalendarById('iotvn30po8f217v5bgbra7p7kg@group.calendar.google.com');

// Create an object from user submission
function Submission(){
  var row = lastRow;
  this.timestamp = sheet.getRange(row, 1).getValue();
  this.name = sheet.getRange(row, 2).getValue();
  this.email = sheet.getRange(row, 3).getValue();
  this.reason = sheet.getRange(row, 4).getValue();
  this.date = sheet.getRange(row, 5).getValue();
  this.time = sheet.getRange(row, 6).getValue();
  this.duration = sheet.getRange(row, 7).getValue();
  this.room = sheet.getRange(row, 8).getValue();
  // Info not from spreadsheet
  this.roomInt = this.room.replace(/^\D+/g, '');
  this.status;
  this.dateString = (this.date.getMonth() + 1) + '/' + this.date.getDate() + '/' + this.date.getYear();
  this.timeString = this.time.toLocaleTimeString();
  this.date.setHours(this.time.getHours());
  this.date.setMinutes(this.time.getMinutes());
  this.calendar = eval('cal' + String(this.roomInt));
  return this;
}

// Use duration to create endTime variable
function getEndTime(request){
  request.endTime = new Date(request.date);
  switch (request.duration){
    case "30 minutes":
      request.endTime.setMinutes(request.date.getMinutes() + 30);
      request.endTimeString = request.endTime.toLocaleTimeString();
      break;
    case "45 minutes":
      request.endTime.setMinutes(request.date.getMinutes() + 45);
      request.endTimeString = request.endTime.toLocaleTimeString();
      break;
    case "1 hour":
      request.endTime.setMinutes(request.date.getMinutes() + 60);
      request.endTimeString = request.endTime.toLocaleTimeString();
      break;
    case "2 hours":
      request.endTime.setMinutes(request.date.getMinutes() + 120);
      request.endTimeString = request.endTime.toLocaleTimeString();
      break;
  }
}

// Check for appointment conflicts
function getConflicts(request){
  var conflicts = request.calendar.getEvents(request.date, request.endTime);
  if (conflicts.length < 1) {
    request.status = "Approve";
  } else {
    request.status = "Conflict";
  }
}

function draftEmail(request){
  request.buttonLink = "https://goo.gl/forms/JX43ZkcyzVKHen4I3";
  request.buttonText = "New Request";
  switch (request.status) {
    case "Approve":
      request.subject = "Confirmation: " + request.room + " Reservation for " + request.dateString;
      request.header = "Confirmation";
      request.message = "Your room reservation has been scheduled.";
      break;
    case "Conflict":
      request.subject = "Conflict with " + request.room + "Reservation for " + request.dateString;
      request.header = "Conflict";
      request.message = "There is a scheduling conflict. Please pick another room or time."
      request.buttonText = "Reschedule";
      break;
  }
}

function updateCalendar(request){
  var event = request.calendar.createEvent(
    request.name,
    request.date,
    request.endTime
    )
}

function sendEmail(request){
  MailApp.sendEmail({
    to: request.email,
    subject: request.header,
    htmlBody: makeEmail(request)
  })
  sheet.getRange(lastRow, lastColumn).setValue("Sent: " + request.status);
}

// --------------- main --------------------

function main(){
  var request = new Submission();
  getEndTime(request);
  getConflicts(request);
  draftEmail(request);
  if (request.status == "Approve") updateCalendar(request);
  sendEmail(request);
}


















