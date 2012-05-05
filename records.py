#!/usr/bin/python

import commands
import re
import shutil
import datetime
import csv
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.message import Message
from email import encoders
from email.utils import COMMASPACE
import mimetypes

completeTotalCalls = 0
completeTotalDuration = 0
avgTalkTime = 0

#List of sales reps
repList = { 'xxxxxxxxxx.csv' : '000000000', 'xxxxxxxxx.csv' : '000000000','xxxxxxx.csv' : '00000000', 'xxxxxxxxx.csv' : '00000000', 'xxxxxxxxx.csv' : '00000000', 'xxxxxxxxx.csv' : '00000000', 'xxxxxxxxxx.csv' : '00000000', 'xxxxxxxx.csv' : '00000000', 'xxxxxxxx.csv':'00000000', 'xxxxxxxxxx.csv' : '000000000', 'xxxxxxxxx.csv' : '00000000', 'xxxxxxxxx.csv' : '000000000' }

#yesterday
#curDate = datetime.datetime.now() - datetime.timedelta(1)

#Get current date
curDate = datetime.datetime.now()
curDate = curDate.strftime("%Y-%m-%d")

for name in repList:

  if (len(str(repList[name])) > 3):
    print "replist[Name] == " + repList[name] + " is greater than 3 in length\n"
    mysql_query_src = 'mysql -u root -e "use asteriskcdrdb; select calldate, src, dst, duration from cdr where calldate like \'' + str(curDate) + '%\' and src=\'' + str(repList[name]) + '\' or calldate like \'' + str(curDate) + '%\' and src=\'' + str(repList[name])[6:] + '\' into outfile \'/tmp/src.' + name + '\'" --password=xxxxxx'
    mysql_query_dst = 'mysql -u root -e "use asteriskcdrdb; select calldate, src, dst, duration from cdr where calldate like \'' + str(curDate) + '%\' and dst=\'' + str(repList[name]) + '\' or calldate like \'' + str(curDate) + '%\' and dst=\'' + str(repList[name])[6:] + '\' into outfile \'/tmp/dst.' + name + '\'" --password=xxxxxx'
  else:
    print "replist[Name] == " + repList[name] + " is less than 3 in length\n"
    mysql_query_src = 'mysql -u root -e "use asteriskcdrdb; select calldate, src, dst, duration from cdr where calldate like \'' + str(curDate) + '%\' and channel like \'%SIP/' + str(repList[name]) + '%\' into outfile \'/tmp/src.' + name + '\'" --password=xxxxxx'
    mysql_query_dst = 'mysql -u root -e "use asteriskcdrdb; select calldate, src, dst, duration from cdr where calldate like \'' + str(curDate) + '%\' and dst=\'' + str(repList[name]) + '\'  into outfile \'/tmp/dst.' + name + '\'" --password=xxxxxxxx'

  #Run sql queries to grab records
  commands.getoutput(mysql_query_src)
  commands.getoutput(mysql_query_dst)

#Fix formatting of CSV files
commands.getoutput("for i in /tmp/*.csv; do sed -i 's/\\s/,/g' $i; done")

#Start reporting file
reportFile = "/tmp/report-" + curDate + ".html"
try:
  report = open(reportFile, "w")
  report.write("<HTML>\n\t<HEAD>\n")
  report.write("\t\t<style type=\"text/css\">\n")
  report.write("\t\t\t\ta { text-decoration: none; }\n")
  report.write("\t\t\t\ta:hover { text-decoration: underline; }\n")
  report.write("\t\t</style>\n")
  report.write("<TITLE> Report for " + curDate + "</TITLE>\n\t</HEAD>\n\t<BODY>\n")

except:
  print "Unable to open report file."

###############################################
# Create bulleted list of names and hyperlinks#
###############################################

report.write("<b><u>Quick links</b></u><BR>")
report.write("<ul style=\"margin: 0px 0px 0px 0px; padding: 5px 5px 5px 10px\">\n")
for name in repList.iterkeys():
  report.write("<li><a href=\"#" + name[:(len(name))-4] + "\">" + name[:(len(name))-4] + "</a><BR></li>\n")
report.write("</ul>\n")

for name in repList:
  duration = 0
  callCount = 0
  completeTotalCalls = 0
  completeTotalDuration = 0

  #read csv to gather average and total time of calls
  timeReader = csv.reader(open("/tmp/src." + name, "rb"), delimiter=',')

  #get total duration and call counts
  for row in timeReader:
    duration += int(row[-1])
    callCount += 1

  #write rep name and column headers
  report.write("\t\t<TABLE style='border: 1px #d79900 solid; text-align: center'><BR>\n")
  report.write("\t\t<TR style='background-color: #CCCCCC'><TD><a name=\"" + name[:(len(name)-4)] + "\"><B> " + name + "</B>  Outgoing Calls</TD></TR><TR><TD>Call Date</TD><TD>Time</TD><TD>Src</TD><TD>Dst</TD><TD>Duration Sec</TD></TR>\n")

  #open rep data call file
  rep = open("/tmp/src." + name, "r")

  #write each call to report file
  for line in rep:
    spLine = line.split(',')
    if (int(spLine[4]) > 20):
      avgTalkTime += int(spLine[4])
      completeTotalCalls += 1

    report.write("\t\t<TR><TD>" + spLine[0] + "</TD><TD>" + spLine[1] + "</TD><TD>" + spLine[2] + "</TD><TD>" + spLine[3] + "</TD><TD>" + spLine[4] + "</TD></TR>\n")

  #close rep data call file
  rep.close()

  if (duration != 0):
    completeTotalDuration = (duration/60)

    report.write("\t\t<TR><TD></TR></TD><TR><TD><B> Total duration: </B></TD><TD>" + str(duration/60) + " min</TD><TD> <B>Avg Call Time:</B></TD><TD>" + str((duration/callCount)) + " sec</TD>")
    if (callCount >= 70):
      report.write("<TD> <B> Total Calls:</B></TD><TD>")
    else:
      report.write("<TD> <B> Total Calls:</B></TD><TD>")
    report.write(str(callCount) + "</TD></TR></TABLE>\n\n")
  else:
    report.write("\t\t<TR><TD> No calls made </TD></TR>")
    completeTotalCalls = 0


  duration = 0
  callCount = 0
  #read csv to gather average and total time of calls
  timeReader = csv.reader(open("/tmp/dst." + name, "rb"), delimiter=',')

  #get total duration and call counts
  for row in timeReader:
    callCount += 1

  #write rep name and column headers
  report.write("\t\t<TABLE style='border: 1px #d79900 solid; text-align: center'><BR>\n")
  report.write("\t\t<TR style='background-color: #CCCCCC'><TD><B> " + name + "</B>  Incoming Calls</TD></TR><TR><TD>Call Date</TD><TD>Time</TD><TD>Src</TD><TD>Dst</TD><TD>Duration Sec</TD></TR>\n")

  #open rep data call file
  rep = open("/tmp/dst." + name, "r")

  #write each call to report file
  for line in rep:
    spLine = line.split(',')
    if (int(spLine[4]) > 20):
     duration += int(row[-1])
     avgTalkTime += int(spLine[4])
     print "Call: "  + spLine[4] +   " greater than 20 sec\n"
     completeTotalCalls += 1

    report.write("\t\t<TR><TD>" + spLine[0] + "</TD><TD>" + spLine[1] + "</TD><TD>" + spLine[2] + "</TD><TD>" + spLine[3] + "</TD><TD>" + spLine[4] + "</TD></TR>\n")

  #close rep data call file
  rep.close()

  if (duration != 0):
    completeTotalDuration += (duration/60)

    report.write("\t\t<TR><TD></TR></TD><TR><TD><B> Total duration: </B></TD><TD>" + str(duration/60) + " min</TD><TD> <B>Avg Call Time:</B></TD><TD>" + str((duration/callCount)) + " sec</TD><TD><B> Total Calls:</B></TD><TD>" + str(callCount) + "</TD></TR>\n\n")
  else:
    report.write("\t\t<TR><TD> No calls made </TD></TR>")

  if (completeTotalCalls >= 70):
    report.write("\t\t<TR><TD></TR></TD><TR><TD style=\"background: green\"><B> Complete Total Calls: " + str(completeTotalCalls) + "</B></TD></TR>")
  else:
    report.write("\t\t<TR><TD></TR></TD><TR><TD style=\"background: red\"><B> Complete Total Calls: " + str(completeTotalCalls) + "</B></TD></TR>")

  if (completeTotalDuration >= 120):
    report.write("\t\t<TR><TD></TR></TD><TR><TD style=\"background: green\"><B> Complete Total Duration: " + str(completeTotalDuration) + " min </B></TD></TR>")
    report.write("\t\t<TR><TD></TR></TD><TR><TD style=\"background: green\"><B> Complete Total Duration: " + str(completeTotalDuration/60) + " hours and " + str(completeTotalDuration % 60) + " min </B></TD></TR>")
  else:
    report.write("\t\t<TR><TD></TR></TD><TR><TD style=\"background: red\"><B> Complete Total Duration: " + str(completeTotalDuration) + " min </B></TD></TR>")
    report.write("\t\t<TR><TD></TR></TD><TR><TD style=\"background: red\"><B> Complete Total Duration: " + str(completeTotalDuration/60) + " hours and " + str(completeTotalDuration % 60) + " min </B></TD></TR>")


report.write("\t</BODY>\n</HTML>")
report.close()

commands.getoutput("/root/removecsv.sh")

try:
  s = smtplib.SMTP('localhost', port=25)
except:
  print "Error connecting to smtp on localhost port 25"

#mail_to = ['xxxxxxxxx@peoplematter.com']
msg_report = MIMEMultipart('alternative')
msg_report['Subject'] = 'Sales phone report for ' + curDate
msg_report['From'] = 'phonereport@peoplematter.com'
msg_report['To'] = COMMASPACE.join(mail_to)
msg = MIMEBase('application', 'octet-stream')
msg.set_payload(file(r'/tmp/report-' + curDate + '.html').read())
encoders.encode_base64(msg)
msg.add_header('Content-Disposition', 'attachment;filename=report-' + curDate + '.html')
msg_report.attach(msg)

try:
  s.sendmail(msg_report['From'], mail_to, msg_report.as_string())
except:
  print "Error sending e-mail."

s.quit()

