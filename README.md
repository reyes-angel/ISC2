# ISC2

Chapter Meetings are held virtually over a popular video conferencing platform in webinar format.   After the meeting, the video conferencing platform provides an attendee report file which is used to determine the number of CPE's that should be allotted to each attendee for attendance.

The attendee report file is a single CSV file containing multiple sections/tables, each is parsed individually and then combined into a single set of data that is enriched and cleansed in order to output the desired result files. 

The sections/tables in the file include:
- Report Generated: General information about the meeting.
- Host Details
- Panelist Details
- Attendee Details
- Other Attended

**Notes**:
- An attendee that joins as a general attendee can be reclassified as a different type of attendee during the meeting.  As an example, somone who joins as an attendee may be re-classified as a panelist while the meeting is in progress.  The user would then have multiple records listed in the attendee report file.
- An attendee may join the call, drop, and then rejoin.   This may happen for many reasons, one being a weak internet connection.  The user would then have multiple records listed in the attendee report file.

In order to properly calculate the earned CPE's the total time attended must be properly calculated.

**The most common use cases include**:
- **Attendee joined and left before the chapter meeting begun**: Official meeting start is at 6:00 PM but the attendee joined and left befor the official meeting start - Discard the record
- **Attendee joined and left after the chapter meeting ended**: Official meeting end is at 8:00 PM but the attendee joined and left after the officla meeting end - Discasrd the record 
- **The attendance interval of one record is completely within the attendance interval of another record**: Attendee joins on two devices, connection #1 goes from 6:00 PM to 6:30 PM. Connection #2 goes from 6:05 PM to 6:10 PM - Discard the record for connection #2 since the time is already accounted for in connection #1
- **The Join Time of one record starts within 60 seconds of another ending**: Attendee joins the call at 6:00 PM as a general attendee and gets reclassified as a panelist at 6:30 PM. The conferencing platform will show a disconnect and then a reconnect a few seconds later. The attendee then remains connected until 8:00 PM. The two records will be "bridged", the attendance time for the attendee will show as going from 6:00 PM to 8:00 PM.  A similar situation occurs for attendees who are disconnected due to a weak internet connection.
- **The Join Time of one record starts more than 60 seconds after another ended**: Attendee joins the call at 6:00 PM and for some reason needs to disconnect to take care of some other task at 6:30.   The other task takes some time to complete but then the attendee rejoins the meeting at 7:30 PM.  We won't bridge the attendance time but we will give the attendee credit for the time they did attend the meeting. 
- **Attendee joined before the meeting started**: Extend join time to beginning of the meeting
- **Attendee left after the meeting started**: Extend leave time to the end of the meeting
      
The default grace period is 10 minutes, if the attendee was present for 110 or more minutes of the meeting they are given credit for having joined the full 120 minute meeting.

A sample attendee report file is included in the sample_file folder.
