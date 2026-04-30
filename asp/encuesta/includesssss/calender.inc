<html><head>
<title>:: AURIS - Calendar ::</title>
<meta http-equiv="Content-Type" content="text/html;charset=utf-8" />
<style type="text/css">
#calendar {
	margin: 0px;
	padding: 0px;
	font-family: Verdana, Tahoma, Arial;
}
#calendar tr.header th {
	font-size: 12px;
	color: #00FFFF;
	text-align: center;
}
#calendar tr.days th {
	font-size: 10px;
	color: #00FFFF;
	text-align: center;
}
#calendar td, #calendar th {
	color: #000000;
	empty-cells: show;
}
#calendar td {
	text-align: right;
	font-size: 12px;
}
#calendar td.weekend {
	background-color: #999999;
}
#calendar td.target {
	background-color: #80c0c0;
}
#calendar a, #calendar a:visited {
	color: #FFFFFF;
	text-decoration: none;
}
#calendar a:hover {
	color: #FF0000;
	text-decoration: none;
}
</style>
<script type="text/javascript">
// Default target date to today.
var targetDate = new Date();

// Initialize the calendar display once the page loads.
window.onload = setCalendar;self.focus();//self.resizeTo(236,195);

//----------------------------------------------------------------------------
// updates the calendar display to reflect the current target date.
//----------------------------------------------------------------------------
function setCalendar(event) {

	var el, tableEl, rowEl, cellEl, linkEl;
	var tmpDate, tmpDate2;
	var i, j;

	// Update month name.
	el = document.getElementById("calendarHeader").firstChild;
	el.nodeValue = targetDate.getMonthName() + String.fromCharCode(32) + targetDate.getFullYear();

	// Start with the first day of the month and go back if necessary to
	// the previous Sunday.
	tmpDate = new Date(Date.parse(targetDate));
	tmpDate.setDate(1);
	while (tmpDate.getDay() != 0) {
		tmpDate.addDays(-1);
	}

	// Go through each calendar day cell in the table and update it.
	tableEl = document.getElementById("calendar");
	for (i = 2; i <= 7; i++) {
		rowEl = tableEl.rows[i];

		// Hide row if it contains no dates in the current month.
		tmpDate2 = new Date(Date.parse(tmpDate));
		tmpDate2.addDays(6);
	/*	if (tmpDate.getMonth() != targetDate.getMonth() &&
				tmpDate2.getMonth() != targetDate.getMonth()) {
			rowEl.style.visibility = "hidden";
			if (document.all)
				for (j = 0; j < rowEl.cells.length; j++)
					rowEl.cells[j].style.borderStyle = "none";
		}
		else {*/
			rowEl.style.visibility = "";
			if (document.all)
				for (j = 0; j < rowEl.cells.length; j++)
					rowEl.cells[j].style.borderStyle = "";
	//	}

		// Loop through a week.
		for (j = 0; j < rowEl.cells.length; j++) {
			cellEl = rowEl.cells[j];
			linkEl = cellEl.firstChild;

			if (tmpDate.getMonth() == targetDate.getMonth()) {
				linkEl.date = new Date(Date.parse(tmpDate));
				s = tmpDate.toString().split(" ");
				linkEl.title = s[0] + " " + s[1] + " " + s[2] + ", " + s[s.length - 1];
				linkEl.firstChild.nodeValue = tmpDate.getDate();
				linkEl.style.visibility = "";
			}
			else
				linkEl.style.visibility = "hidden";

			// Highlight the current date.
			if (cellEl.oldClass == null)
				cellEl.oldClass = cellEl.className;
			if (Date.parse(tmpDate) == Date.parse(targetDate))
				cellEl.className = cellEl.oldClass + " target";
			else
				cellEl.className = cellEl.oldClass;

			// Go to the next day.
			tmpDate.addDays(1);
		}
	}
}

//----------------------------------------------------------------------------
// Event handlers for the calendar elements.
//----------------------------------------------------------------------------
function addMonths(event, n) {

	// Advance the calendar month and update the display.
	targetDate.addMonths(n);
	setCalendar(event);
}

function addYears(event, n) {

	// Advance the calendar year and update the display.
	targetDate.addYears(n);
	setCalendar(event);
}

function setTargetDate(event, link) {

	var strURL;

	// Change the target date and update the calendar.
	if (link.date != null) {
		targetDate = new Date(Date.parse(link.date));
		setCalendar(event);
		//displayDate(event);
		strURL = self.location.toString();
		eval("self.opener."+strURL.substr(strURL.indexOf("?")+1)+".value = formatDate(targetDate)");
	}
}

function displayDate(event) {

	// Display the current target date as a formatted string.
	alert(formatDate(targetDate));
}

function formatDate() {

	var mm, dd, yyyy;

	// Return the target date in "mm/dd/yyyy" format.
	mm = String(targetDate.getMonth() + 1);
	while (mm.length < 2)
		mm = "0" + mm;
	dd = String(targetDate.getDate());
	while (dd.length < 2)
		dd = "0" + dd;
	yyyy = String(targetDate.getFullYear());
	while (yyyy.length < 4)
		yyyy = "0" + yyyy;

	return mm + "/" + dd + "/" + yyyy;
}

//============================================================================
// Add new properties and methods to the Date object.
//============================================================================

// Properties
Date.prototype.monthNames = new Array("January", "February", "March"+String.fromCharCode(32,32,32,32), "April"+String.fromCharCode(32,32,32,32), "May"+String.fromCharCode(32,32,32,32), "June"+String.fromCharCode(32,32,32,32), "July"+String.fromCharCode(32,32,32,32), "August", "September", "October", "November", "December");
Date.prototype.savedDate  = null;

// Methods
Date.prototype.getMonthName = dateGetMonthName;
Date.prototype.getDays		= dateGetDays;
Date.prototype.addDays		= dateAddDays;
Date.prototype.addMonths	= dateAddMonths;
Date.prototype.addYears		= dateAddYears;

//----------------------------------------------------------------------------
// getMonthName(): Returns the name of the date's month.
//----------------------------------------------------------------------------
function dateGetMonthName() {

	return this.monthNames[this.getMonth()];
}

//----------------------------------------------------------------------------
// getDays(): Returns the number of days in the date's month.
//----------------------------------------------------------------------------
function dateGetDays() {

	var tmpDate, d, m;

	tmpDate = new Date(Date.parse(this));
	m = tmpDate.getMonth();
	d = 28;
	do {
		d++;
		tmpDate.setDate(d);
	} while (tmpDate.getMonth() == m);

	return d - 1;
}

//----------------------------------------------------------------------------
// addDays(n): Adds the specified number of days to the date.
//----------------------------------------------------------------------------
function dateAddDays(n) {

	// Add the specified number of days.
	this.setDate(this.getDate() + n);

	// Reset the new day of month.
	this.savedDate = this.getDate();
}

//----------------------------------------------------------------------------
// addMonths(n): Adds the specified number of months to the date, adjusting
// the day of the month if necessary.
//----------------------------------------------------------------------------
function dateAddMonths(n) {

	// Save the day of month if not already set.
	if (this.savedDate == null)
		this.savedDate = this.getDate();

	// Set the day of month to the first to avoid rolling.
	this.setDate(1);

	// Add the specified number of months.
	this.setMonth(this.getMonth() + n);

	// Restore the saved day of month, if possible.
	this.setDate(Math.min(this.savedDate, this.getDays()));
}

//----------------------------------------------------------------------------
// addYears(n): Adds the specified number of years to the date, adjusting the
// day of the month for leap years.
//----------------------------------------------------------------------------
function dateAddYears(n) {

	// Save the day of month if not already set.
	if (this.savedDate == null)
		this.savedDate = this.getDate();

	// Set the day of month to the first to avoid rolling.
	this.setDate(1);

	// Add the specified number of years.
	this.setFullYear(this.getFullYear() + n);

	// Restore the saved day of month, if possible.
	this.setDate(Math.min(this.savedDate, this.getDays()));
}
</script>
</head>
<body bgcolor="#00496B" topmargin="1" leftmargin="1" marginheight="1" marginwidth="1">

<table id="calendar" border="1" bordercolor="#0099CC" cellpadding="2" cellpadding="2" width="100%">
<tr class="header">
	<th background="../images/bg-middlebar.gif"><a href="" onclick="addYears(event, -1); return false;"><img src='../buttons/calendar-lastyear.gif' alt='Goto Last Year' name='last-year' width=27 height=20 border=0></a></th>
	<th background="../images/bg-middlebar.gif"><a href="" onclick="addMonths(event, -1); return false;"><img src='../buttons/calendar-lastmonth.gif' alt='Goto Last Month' name='last-month' width=27 height=20 border=0></a></th>
	<th id="calendarHeader" colspan="3" background="../images/bg-middlebar.gif">&nbsp;</th>
	<th background="../images/bg-middlebar.gif"><a href="" onclick="addMonths(event, 1); return false;"><img src='../buttons/calendar-month-rt.gif' alt='Goto Next Month' name='next-month' width=27 height=20 border=0></a></th>
	<th background="../images/bg-middlebar.gif"><a href="" onclick="addYears(event, 1); return false;"><img src='../buttons/calendar-nextyr.gif' alt='Goto Next Year' name='next-year' width=27 height=20 border=0></a></th>
</tr>
<tr class="days">
	<th>Sun</th>
	<th>Mon</th>
	<th>Tue</th>
	<th>Wed</th>
	<th>Thu</th>
	<th>Fri</th>
	<th>Sat</th>
</tr>
<tr>
	<td class="weekend"><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
	<td><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
	<td><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
	<td><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
	<td><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
	<td><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
	<td class="weekend"><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
</tr>
<tr>
	<td class="weekend"><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
	<td><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
	<td><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
	<td><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
	<td><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
	<td><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
	<td class="weekend"><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
</tr>
<tr>
	<td class="weekend"><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
	<td><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
	<td><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
	<td><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
	<td><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
	<td><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
	<td class="weekend"><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
</tr>
<tr>
	<td class="weekend"><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
	<td><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
	<td><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
	<td><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
	<td><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
	<td><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
	<td class="weekend"><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
</tr>
<tr>
	<td class="weekend"><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
	<td><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
	<td><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
	<td><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
	<td><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
	<td><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
	<td class="weekend"><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
</tr>
<tr>
	<td class="weekend"><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
	<td><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
	<td><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
	<td><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
	<td><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
	<td><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
	<td class="weekend"><a href="" onclick="setTargetDate(event, this); return false;">&nbsp;</a></td>
</tr>
</table>
</body>
</html>