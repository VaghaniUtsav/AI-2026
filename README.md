# AI-2026

-- What report did we drill into?
Sel Report Name = SELECTEDVALUE('Report Display'[Report Name], "All Reports")


Rpt Unique Users =
CALCULATE(
    DISTINCTCOUNT('SESSION USAGE'[UserName]),
    TREATAS( VALUES('Report Display'[Report Name]), 'REPORT CLASSIFICATION'[Report Name] )
)

Rpt Sessions =
CALCULATE(
    COUNT('SESSION USAGE'[UsageId]),
    TREATAS( VALUES('Report Display'[Report Name]), 'REPORT CLASSIFICATION'[Report Name] )
)

Rpt Active Days =
CALCULATE(
    DISTINCTCOUNT('SESSION USAGE'[Date]),
    TREATAS( VALUES('Report Display'[Report Name]), 'REPORT CLASSIFICATION'[Report Name] )
)

Rpt Avg Sessions per User = DIVIDE( [Rpt Sessions], [Rpt Unique Users] )

-- This report's share of the whole platform (sessions)
Rpt Share of Platform % =
VAR rpt = [Rpt Sessions]
VAR platform =
    CALCULATE( COUNT('SESSION USAGE'[UsageId]),
        REMOVEFILTERS('Report Display'),
        KEEPFILTERS('REPORT CLASSIFICATION'[Report Name] = "Total") )
RETURN DIVIDE( rpt, platform )


-- Power users: 5+ sessions of this report
Rpt Power Users (5+) =
COUNTROWS( FILTER( VALUES('SESSION USAGE'[UserName]), [User Sessions] >= 5 ) )

-- One-time users: used it exactly once
Rpt One-time Users =
COUNTROWS( FILTER( VALUES('SESSION USAGE'[UserName]), [User Sessions] = 1 ) )

-- % of this report's users who are power users
Rpt Power User % = DIVIDE( [Rpt Power Users (5+)], [Rpt Unique Users] )

-- New users this period (first-ever session of THIS report falls in the current window).
-- Uses _Period Start/_Latest Data Date from MyBI_LikeForLike_Measures.md.
Rpt New Users =
VAR S = [_Period Start]
VAR E = [_Period End]
RETURN
    CALCULATE(
        DISTINCTCOUNT('SESSION USAGE'[UserName]),
        TREATAS( VALUES('Report Display'[Report Name]), 'REPORT CLASSIFICATION'[Report Name] ),
        FILTER(
            VALUES('SESSION USAGE'[UserName]),
            VAR fs =
                CALCULATE(
                    MIN('SESSION USAGE'[Date]),
                    TREATAS( VALUES('Report Display'[Report Name]), 'REPORT CLASSIFICATION'[Report Name] ),
                    ALLEXCEPT('SESSION USAGE', 'SESSION USAGE'[UserName])
                )
            RETURN fs >= S && fs <= E
        )
    )

Rpt Returning Users = [Rpt Unique Users] - [Rpt New Users]


User Sessions =
CALCULATE(
    COUNT('SESSION USAGE'[UsageId]),
    TREATAS( VALUES('Report Display'[Report Name]), 'REPORT CLASSIFICATION'[Report Name] )
)

User Active Days =
CALCULATE(
    DISTINCTCOUNT('SESSION USAGE'[Date]),
    TREATAS( VALUES('Report Display'[Report Name]), 'REPORT CLASSIFICATION'[Report Name] )
)

User First Seen =
CALCULATE(
    MIN('SESSION USAGE'[Date]),
    TREATAS( VALUES('Report Display'[Report Name]), 'REPORT CLASSIFICATION'[Report Name] )
)

User Last Seen =
CALCULATE(
    MAX('SESSION USAGE'[Date]),
    TREATAS( VALUES('Report Display'[Report Name]), 'REPORT CLASSIFICATION'[Report Name] )
)

-- Days since the user last opened this report (recency). Uses _Latest Data Date.
User Days Since Last =
VAR last = [User Last Seen]
RETURN IF( ISBLANK(last), BLANK(), INT( [_Latest Data Date] - last ) )

-- How long the user has been active on this report (tenure)
User Tenure Days =
VAR first = [User First Seen]
RETURN IF( ISBLANK(first), BLANK(), INT( [_Latest Data Date] - first ) )

-- This user's share of the report's sessions
User Share of Report % = DIVIDE( [User Sessions], [Rpt Sessions] )

-- Intensity: sessions per active day
User Sessions per Active Day = DIVIDE( [User Sessions], [User Active Days] )

-- Rough visit cadence: days between first and last, spread over visits
User Avg Days Between Visits =
VAR d = [User Active Days]
RETURN IF( d <= 1, BLANK(), DIVIDE( [User Tenure Days], d - 1 ) )

-- Rank this user within the report (1 = heaviest user)
User Rank in Report =
IF( NOT ISBLANK([User Sessions]),
    RANKX( ALLSELECTED('SESSION USAGE'[UserName]), [User Sessions], , DESC, DENSE ) )


    Rpt Users by Month =
CALCULATE(
    DISTINCTCOUNT('SESSION USAGE'[UserName]),
    TREATAS( VALUES('Report Display'[Report Name]), 'REPORT CLASSIFICATION'[Report Name] )
)

Rpt Sessions by Month =
CALCULATE(
    COUNT('SESSION USAGE'[UsageId]),
    TREATAS( VALUES('Report Display'[Report Name]), 'REPORT CLASSIFICATION'[Report Name] )
)

Users • 1 session      = COUNTROWS( FILTER( VALUES('SESSION USAGE'[UserName]), [User Sessions] = 1 ) )
Users • 2-4 sessions   = COUNTROWS( FILTER( VALUES('SESSION USAGE'[UserName]), [User Sessions] >= 2  && [User Sessions] <= 4 ) )
Users • 5-9 sessions   = COUNTROWS( FILTER( VALUES('SESSION USAGE'[UserName]), [User Sessions] >= 5  && [User Sessions] <= 9 ) )
Users • 10-19 sessions = COUNTROWS( FILTER( VALUES('SESSION USAGE'[UserName]), [User Sessions] >= 10 && [User Sessions] <= 19 ) )
Users • 20+ sessions   = COUNTROWS( FILTER( VALUES('SESSION USAGE'[UserName]), [User Sessions] >= 20 ) )
