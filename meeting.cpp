#include "StdAfx.h"
#include "meeting.h"

//Subject
CString meeting::GetSubject() const
{
	return Subject;
}

void meeting::SetSubject(const CString& subject)
{
	Subject = subject;
}

//DateAndTime
void meeting::SetCurrentTime()
{
	COleDateTime DateAndTime = COleDateTime::GetCurrentTime();
	SetDate(DateAndTime.GetYear(), DateAndTime.GetMonth(), DateAndTime.GetDay());
	SetTime(DateAndTime.GetHour(), DateAndTime.GetMinute(), DateAndTime.GetSecond());
}

int meeting::GetYear() const
{
	return Year;
}

void meeting::SetYear(const int& year)
{
	Year = year;
}

int meeting::GetMonth() const
{
	return Month;
}

void meeting::SetMonth(const int& month)
{
	Month = month;
}

void meeting::SetDay(const int& day)
{
	Day = day;
}

void meeting::SetDate(const int& year, const int& month, const int& day)
{
	SetYear(year);
	SetMonth(month);
	SetDay(day);
}

int meeting::GetHour() const
{
	return Hour;
}

void meeting::SetHour(const int& hur)
{
	Hour = hur;
}

int meeting::GetMin() const
{
	return Minute;
}

void meeting::SetMin(const int& min)
{
	Minute = min;
}

int meeting::GetSec() const
{
	return Second;
}

void meeting::SetSec(const int& sec)
{
	Second = sec;
}

void meeting::SetTime(const int& hur, const int& min, const int& sec)
{
	SetHour(hur);
	SetMin(min);
	SetSec(sec);
}

COleDateTime meeting::GetDateAndTime() const
{
	COleDateTime DateAndTime(Year, Month, Day, Hour, Minute, Second);
	return DateAndTime;
}

//Place
CString meeting::GetPlace() const
{
	return Place;
}

void meeting::SetPlace(const CString& place)
{
	Place = place;
}

//nReminderMinutesBeforeStart
long meeting::GetMinutesBeforeStart() const
{
	return nReminderMinutesBeforeStart;
}

void meeting::SetHowLongBeforeStart(const long& minutes)
{
	nReminderMinutesBeforeStart = minutes;
}

//nFewMinutes;
long meeting::GetMinutesActive() const
{
	return nFewMinutes;
}

void meeting::SetHowlongActive(const long& minutes)
{
	nFewMinutes = minutes;
}

//lpszMeetingDescription
CString meeting::GetDescription() const
{
	return lpszMeetingDescription;
}

void meeting::SetDescription(const CString& description)
{
	lpszMeetingDescription = description;
}
