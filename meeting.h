#include<vector>

class meeting
{
private:
	CString Subject;
public:
	CString GetSubject() const;
	void    SetSubject(const CString& subject);

private:
	int Year, Month, Day, Hour, Minute, Second;
public:
	void SetCurrentTime();

	int GetYear() const;   void SetYear(const int& year);
	int GetMonth() const;  void SetMonth(const int& month);
	int GetDay() const;    void SetDay(const int& day);
	void SetDate(const int& year, const int& month, const int& day);

	int GetHour() const;   void SetHour(const int& hur);
	int GetMin() const;    void SetMin(const int& min);
	int GetSec() const;    void SetSec(const int& sec);
	void SetTime(const int& hur, const int& min, const int& sec);

	COleDateTime GetDateAndTime() const;

private:
	CString Place;
public:
	CString GetPlace() const;
	void    SetPlace(const CString& location);

private:
	long    nReminderMinutesBeforeStart;
public:
	long GetMinutesBeforeStart() const;
	void SetHowLongBeforeStart(const long& minutes);

private:
	long    nFewMinutes;
public:
	long GetMinutesActive() const;
	void SetHowlongActive(const long& minutes);

private:
	CString lpszMeetingDescription;
public:
	CString GetDescription() const;
	void    SetDescription(const CString& description);
};