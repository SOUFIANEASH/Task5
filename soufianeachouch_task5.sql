-- =========================================================
-- Practical Assignment 5
-- Author: Soufiane Achouch
-- Database: CRA (Mobile Network Operators) - Microsoft SQL Server (T-SQL)
--
-- Requirements
--   * Assignment 4 must have been run first (Person.email, Subscriber.discount,
--     Subscriber.mobTelID are used here).
--   * Run on your OWN copy of CRA: creating views/functions/procedures needs
--     CREATE rights, and the school server's CRA is shared by the whole class.
--   * CREATE OR ALTER (SQL Server 2016 SP1+) makes the script re-runnable.
--
-- Conventions
--   * Service IDs (table Service): 1 = calls to Lithuanian networks,
--     2 = SMS to Lithuanian networks, 3 = MMS, 4 = calls to foreign networks,
--     5 = SMS to foreign networks, 6 = video call.
--     "Call" = service 1, 4 or 6.  "SMS" = service 2 or 5.  "MMS" = service 3.
--   * Duration of a call = DATEDIFF(SECOND, connectionStart, connectionEnd).
-- =========================================================

-- SAFETY: stops the script on the shared school server (see Assignment 4).
IF @@SERVERNAME LIKE 'MSQLDBLAB%'
BEGIN
    PRINT 'Stopped: shared school server. Run this on your own local copy of CRA.';
    SET NOEXEC ON;   -- run SET NOEXEC OFF; to undo
END
GO

USE CRA;
GO


-- ---------------------------------------------------------
-- 1. Check that the email address of every person is correctly formatted
--    A valid email: exactly one '@', only letters, digits and . _ - characters,
--    a non-empty local part, a dot inside the domain, no leading/trailing dot,
--    no two dots in a row.
--    Column "matchesAssignment4Format" additionally checks the format required
--    in Assignment 4: firstname.lastname@<one of the Domain table domains>.
--    The binary collation makes the character-range test exact (diacritics,
--    spaces and apostrophes are rejected).
-- ---------------------------------------------------------
SELECT
    p.personID,
    p.firstName,
    p.lastName,
    p.email,
    CASE
        WHEN p.email IS NULL THEN 'Missing'
        WHEN v.atCount = 1
         AND v.atPos > 1
         AND p.email COLLATE Latin1_General_BIN NOT LIKE '%[^a-zA-Z0-9._@-]%'
         AND CHARINDEX('.', p.email, v.atPos) > v.atPos + 1
         AND LEFT(p.email, 1)  <> '.'
         AND RIGHT(p.email, 1) <> '.'
         AND p.email NOT LIKE '%..%'
         AND p.email NOT LIKE '%.@%'
         AND p.email NOT LIKE '%@.%'
        THEN 'Valid'
        ELSE 'Invalid'
    END AS emailStatus,
    CASE WHEN EXISTS (
            SELECT 1
            FROM dbo.[Domain] d
            WHERE p.email = LOWER(p.firstName) + '.' + LOWER(p.lastName) + '@' + d.domainName)
         THEN 'Yes' ELSE 'No' END AS matchesAssignment4Format
FROM dbo.Person p
CROSS APPLY (
    SELECT LEN(p.email) - LEN(REPLACE(p.email, '@', '')) AS atCount,
           CHARINDEX('@', p.email)                        AS atPos
) v
ORDER BY emailStatus DESC, p.personID;

-- Summary of the same check (how many valid / invalid / missing)
WITH EmailCheck AS (
    SELECT
        CASE
            WHEN p.email IS NULL THEN 'Missing'
            WHEN v.atCount = 1
             AND v.atPos > 1
             AND p.email COLLATE Latin1_General_BIN NOT LIKE '%[^a-zA-Z0-9._@-]%'
             AND CHARINDEX('.', p.email, v.atPos) > v.atPos + 1
             AND LEFT(p.email, 1)  <> '.'
             AND RIGHT(p.email, 1) <> '.'
             AND p.email NOT LIKE '%..%'
             AND p.email NOT LIKE '%.@%'
             AND p.email NOT LIKE '%@.%'
            THEN 'Valid'
            ELSE 'Invalid'
        END AS emailStatus
    FROM dbo.Person p
    CROSS APPLY (
        SELECT LEN(p.email) - LEN(REPLACE(p.email, '@', '')) AS atCount,
               CHARINDEX('@', p.email)                        AS atPos
    ) v
)
SELECT emailStatus, COUNT(*) AS persons
FROM EmailCheck
GROUP BY emailStatus;
GO


-- ---------------------------------------------------------
-- 2. Popularity of email service providers (the part after '@')
--    Persons without a usable email are ignored; the percentage is
--    calculated from the persons that have one.
-- ---------------------------------------------------------
SELECT
    SUBSTRING(email, CHARINDEX('@', email) + 1, LEN(email)) AS emailProvider,
    COUNT(*)                                                AS persons,
    CAST(100.0 * COUNT(*) / SUM(COUNT(*)) OVER () AS DECIMAL(5,2)) AS percentage
FROM dbo.Person
WHERE email LIKE '%_@_%'
GROUP BY SUBSTRING(email, CHARINDEX('@', email) + 1, LEN(email))
ORDER BY persons DESC, emailProvider;
GO


-- ---------------------------------------------------------
-- 3. View: total call duration of each subscriber
--    All subscribers are listed (LEFT JOIN): those without calls get 0.
--    SMS (services 2 and 5) are not calls and are excluded.
-- ---------------------------------------------------------
CREATE OR ALTER VIEW dbo.vw_SubscriberTotalCallDuration
AS
SELECT
    s.subscriberID,
    s.phoneNumber,
    pe.firstName,
    pe.lastName,
    ISNULL(c.callCount, 0) AS callCount,
    t.totalSeconds         AS totalCallSeconds,
    CONCAT(t.totalSeconds / 3600, ':',
           RIGHT('0' + CAST(t.totalSeconds % 3600 / 60 AS VARCHAR(2)), 2), ':',
           RIGHT('0' + CAST(t.totalSeconds % 60 AS VARCHAR(2)), 2)) AS totalCallDuration   -- hh:mm:ss
FROM dbo.Subscriber s
JOIN dbo.Person pe ON pe.personID = s.personID
LEFT JOIN (
    SELECT subscriberID,
           COUNT(*) AS callCount,
           SUM(CAST(DATEDIFF(SECOND, connectionStart, connectionEnd) AS BIGINT)) AS totalSeconds
    FROM dbo.P2PConnection
    WHERE serviceID IN (1, 4, 6)
    GROUP BY subscriberID
) c ON c.subscriberID = s.subscriberID
CROSS APPLY (SELECT ISNULL(c.totalSeconds, 0) AS totalSeconds) t;
GO


-- ---------------------------------------------------------
-- 4. View: the longest call of each person and who received it
--    A person may have several contracts, so calls of all their subscriber
--    records are compared. If two calls are equally long, the earlier one
--    is shown. The recipient is identified by number; if that number belongs
--    to a subscriber of this database, the recipient's name is shown too.
--    A NULL recipient country means a call inside Lithuania.
-- ---------------------------------------------------------
CREATE OR ALTER VIEW dbo.vw_PersonLongestCall
AS
WITH CallDurations AS (
    SELECT
        s.personID,
        c.subscriberID,
        c.recipientNumber,
        c.recipientCountry,
        c.connectionStart,
        c.connectionEnd,
        DATEDIFF(SECOND, c.connectionStart, c.connectionEnd) AS durationSeconds
    FROM dbo.P2PConnection c
    JOIN dbo.Subscriber s ON s.subscriberID = c.subscriberID
    WHERE c.serviceID IN (1, 4, 6)
),
Ranked AS (
    SELECT *,
           ROW_NUMBER() OVER (PARTITION BY personID
                              ORDER BY durationSeconds DESC, connectionStart) AS rn
    FROM CallDurations
)
SELECT
    pe.personID,
    pe.firstName,
    pe.lastName,
    r.subscriberID                       AS callerSubscriberID,
    r.connectionStart,
    r.durationSeconds,
    CONCAT(r.durationSeconds / 3600, ':',
           RIGHT('0' + CAST(r.durationSeconds % 3600 / 60 AS VARCHAR(2)), 2), ':',
           RIGHT('0' + CAST(r.durationSeconds % 60 AS VARCHAR(2)), 2)) AS duration,   -- hh:mm:ss
    r.recipientNumber,
    rec.firstName                        AS recipientFirstName,
    rec.lastName                         AS recipientLastName,
    COALESCE(cn.countryName, 'Lithuania') AS recipientCountry
FROM Ranked r
JOIN dbo.Person pe ON pe.personID = r.personID
OUTER APPLY (
    SELECT TOP (1) rp.firstName, rp.lastName
    FROM dbo.Subscriber rs
    JOIN dbo.Person rp ON rp.personID = rs.personID
    WHERE rs.phoneNumber = r.recipientNumber
    ORDER BY rs.contractStartDate DESC
) rec
LEFT JOIN dbo.Country cn ON cn.countryCode = r.recipientCountry
WHERE r.rn = 1;
GO


-- ---------------------------------------------------------
-- 5. View: for each mobile network operator, the subscriber who transferred
--    the largest amount of internet data (uploaded + downloaded = amountMB
--    summed over all internet connections). The operator is the operator of
--    the subscriber's plan. One subscriber per operator (ties: lowest ID).
-- ---------------------------------------------------------
CREATE OR ALTER VIEW dbo.vw_TopDataSubscriberPerOperator
AS
WITH DataUsage AS (
    SELECT subscriberID, SUM(amountMB) AS totalMB
    FROM dbo.InternetConnection
    GROUP BY subscriberID
),
Ranked AS (
    SELECT
        p.operatorID,
        s.subscriberID,
        s.personID,
        s.phoneNumber,
        u.totalMB,
        ROW_NUMBER() OVER (PARTITION BY p.operatorID
                           ORDER BY u.totalMB DESC, s.subscriberID) AS rn
    FROM DataUsage u
    JOIN dbo.Subscriber s ON s.subscriberID = u.subscriberID
    JOIN dbo.[Plan] p     ON p.planID = s.planID
)
SELECT
    RTRIM(mno.operatorName) AS operatorName,
    r.subscriberID,
    pe.firstName,
    pe.lastName,
    r.phoneNumber,
    r.totalMB
FROM Ranked r
JOIN dbo.MobileNetworkOperator mno ON mno.operatorID = r.operatorID
JOIN dbo.Person pe                 ON pe.personID = r.personID
WHERE r.rn = 1;
GO


-- ---------------------------------------------------------
-- 6. Function: amount payable for LOCAL calls of a subscriber in a month
--    Local call = service 1 (calls to Lithuanian networks) made in Lithuania
--    (connectionInitializationPlace IS NULL).
--    * The price per minute is PlanBundle.price of the subscriber's plan for
--      service 1 (e.g. plan 1 = 0.20 Lt per minute, plans 2 and 3 = 0.00).
--    * Calls are charged per second (duration in minutes = seconds / 60.0);
--      to charge every started minute use CEILING(seconds / 60.0).
--    Free minutes / discounts are not part of this function (see 9).
-- ---------------------------------------------------------
CREATE OR ALTER FUNCTION dbo.fn_LocalCallsAmount
(
    @subscriberID INT,
    @year         INT,
    @month        INT
)
RETURNS DECIMAL(12,2)
AS
BEGIN
    DECLARE @monthStart DATETIME = DATEFROMPARTS(@year, @month, 1);
    DECLARE @monthEnd   DATETIME = DATEADD(MONTH, 1, @monthStart);
    DECLARE @amount DECIMAL(12,2);

    SELECT @amount = SUM(DATEDIFF(SECOND, c.connectionStart, c.connectionEnd) / 60.0 * pb.price)
    FROM dbo.P2PConnection c
    JOIN dbo.Subscriber s  ON s.subscriberID = c.subscriberID
    JOIN dbo.PlanBundle pb ON pb.planID = s.planID
                          AND pb.serviceID = c.serviceID
    WHERE c.subscriberID = @subscriberID
      AND c.serviceID = 1
      AND c.connectionInitializationPlace IS NULL
      AND c.connectionStart >= @monthStart
      AND c.connectionStart <  @monthEnd;

    SET @amount = ISNULL(@amount, 0);   -- no calls in that month = nothing to pay
    RETURN @amount;
END;
GO


-- ---------------------------------------------------------
-- 7. Function: pairs of subscribers who communicated with each other in a
--    month. "Communicated" = a call or an SMS (services 1, 2, 4, 5, 6) to the
--    other subscriber's phone number. Bidirectional: A contacted B AND B
--    contacted A in that month. Each pair is returned once (ID1 < ID2).
-- ---------------------------------------------------------
CREATE OR ALTER FUNCTION dbo.fn_MutualContacts
(
    @year  INT,
    @month INT
)
RETURNS TABLE
AS
RETURN
    WITH Contacts AS (
        SELECT DISTINCT
               c.subscriberID  AS fromSubscriberID,
               rs.subscriberID AS toSubscriberID
        FROM dbo.P2PConnection c
        JOIN dbo.Subscriber rs ON rs.phoneNumber = c.recipientNumber
        WHERE c.serviceID IN (1, 2, 4, 5, 6)
          AND c.connectionStart >= DATEFROMPARTS(@year, @month, 1)
          AND c.connectionStart <  DATEADD(MONTH, 1, DATEFROMPARTS(@year, @month, 1))
          AND c.subscriberID <> rs.subscriberID
    )
    SELECT
        a.fromSubscriberID                  AS subscriber1ID,
        s1.phoneNumber                      AS phone1,
        p1.firstName + ' ' + p1.lastName    AS name1,
        a.toSubscriberID                    AS subscriber2ID,
        s2.phoneNumber                      AS phone2,
        p2.firstName + ' ' + p2.lastName    AS name2
    FROM Contacts a
    JOIN Contacts b      ON b.fromSubscriberID = a.toSubscriberID
                        AND b.toSubscriberID   = a.fromSubscriberID
    JOIN dbo.Subscriber s1 ON s1.subscriberID = a.fromSubscriberID
    JOIN dbo.Person p1     ON p1.personID = s1.personID
    JOIN dbo.Subscriber s2 ON s2.subscriberID = a.toSubscriberID
    JOIN dbo.Person p2     ON p2.personID = s2.personID
    WHERE a.fromSubscriberID < a.toSubscriberID;
GO


-- ---------------------------------------------------------
-- 8. Stored procedure: renew a subscriber's contract under the same conditions
--    A renewal is a NEW contract (new Subscriber row) that
--      * starts the day after the old contract ended and lasts as long as it
--        (same number of days),
--      * keeps the person, plan, phone number, discount and mobile phone,
--      * copies the additional services, internet data plans and discount
--        packages that were still active when the old contract ended.
--    It fails if the subscriber does not exist, has no contract end date, or
--    the contract was already renewed. All changes happen in one transaction.
--    Returns the ID of the new contract in @newSubscriberID.
-- ---------------------------------------------------------
CREATE OR ALTER PROCEDURE dbo.sp_RenewContract
    @subscriberID    INT,
    @newSubscriberID INT = NULL OUTPUT
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;

    DECLARE @start DATE, @end DATE, @phoneNumber CHAR(10);
    DECLARE @newStart DATE, @newEnd DATE;
    DECLARE @endNextDay DATETIME;

    SELECT @start = contractStartDate,
           @end = contractEndDate,
           @phoneNumber = phoneNumber
    FROM dbo.Subscriber
    WHERE subscriberID = @subscriberID;

    IF @start IS NULL
        THROW 50001, 'Subscriber not found.', 1;
    IF @end IS NULL
        THROW 50002, 'The contract has no end date, so there is nothing to renew.', 1;

    SET @newStart   = DATEADD(DAY, 1, @end);
    SET @newEnd     = DATEADD(DAY, DATEDIFF(DAY, @start, @end), @newStart);
    SET @endNextDay = DATEADD(DAY, 1, CAST(@end AS DATETIME));   -- first moment after the old contract

    IF EXISTS (SELECT 1 FROM dbo.Subscriber
               WHERE phoneNumber = @phoneNumber AND contractStartDate >= @newStart)
        THROW 50003, 'This contract has already been renewed.', 1;

    BEGIN TRY
        BEGIN TRANSACTION;

        IF COLUMNPROPERTY(OBJECT_ID('dbo.Subscriber'), 'subscriberID', 'IsIdentity') = 1
        BEGIN
            INSERT INTO dbo.Subscriber
                   (personID, planID, phoneNumber, contractStartDate, contractEndDate, discount, mobTelID)
            SELECT personID, planID, phoneNumber, @newStart, @newEnd, discount, mobTelID
            FROM dbo.Subscriber
            WHERE subscriberID = @subscriberID;

            SET @newSubscriberID = SCOPE_IDENTITY();
        END
        ELSE
        BEGIN
            SELECT @newSubscriberID = ISNULL(MAX(subscriberID), 0) + 1
            FROM dbo.Subscriber WITH (UPDLOCK, HOLDLOCK);

            INSERT INTO dbo.Subscriber
                   (subscriberID, personID, planID, phoneNumber, contractStartDate, contractEndDate, discount, mobTelID)
            SELECT @newSubscriberID, personID, planID, phoneNumber, @newStart, @newEnd, discount, mobTelID
            FROM dbo.Subscriber
            WHERE subscriberID = @subscriberID;
        END

        -- additional services active at the end of the old contract
        INSERT INTO dbo.OrderedAdditionalService
               (subscriberID, operatorAdditionalServiceID, orderDate, cancellationDate)
        SELECT @newSubscriberID, o.operatorAdditionalServiceID, CAST(@newStart AS DATETIME), NULL
        FROM dbo.OrderedAdditionalService o
        WHERE o.subscriberID = @subscriberID
          AND o.orderDate < @endNextDay
          AND (o.cancellationDate IS NULL OR o.cancellationDate >= @endNextDay);

        -- internet data plans active at the end of the old contract
        INSERT INTO dbo.OrderedInternetDataPlan
               (internetDataPlanID, subscriberID, orderDate, cancellationDate)
        SELECT o.internetDataPlanID, @newSubscriberID, CAST(@newStart AS DATETIME), NULL
        FROM dbo.OrderedInternetDataPlan o
        WHERE o.subscriberID = @subscriberID
          AND o.orderDate < @endNextDay
          AND (o.cancellationDate IS NULL OR o.cancellationDate >= @endNextDay);

        -- discount packages active at the end of the old contract
        INSERT INTO dbo.OrderedDiscount
               (subscriberID, packageID, orderDate, cancellationDate)
        SELECT @newSubscriberID, o.packageID, CAST(@newStart AS DATETIME), NULL
        FROM dbo.OrderedDiscount o
        WHERE o.subscriberID = @subscriberID
          AND o.orderDate < @endNextDay
          AND (o.cancellationDate IS NULL OR o.cancellationDate >= @endNextDay);

        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;
        THROW;
    END CATCH;

    SELECT * FROM dbo.Subscriber WHERE subscriberID = @newSubscriberID;   -- the new contract
END;
GO


-- ---------------------------------------------------------
-- 9a. Helper: which part of a month an item was valid
--     Returns a number between 0 and 1: the number of days of the given month
--     in which [@activeFrom, @activeTo] was valid (both days included), divided
--     by the number of days in the month. @activeTo = NULL means "still valid".
--     Examples: valid for 15 of 30 days = 0.5, for a quarter of the month = 0.25.
-- ---------------------------------------------------------
CREATE OR ALTER FUNCTION dbo.fn_ActiveMonthFraction
(
    @activeFrom DATETIME,
    @activeTo   DATETIME,
    @year       INT,
    @month      INT
)
RETURNS DECIMAL(9,6)
AS
BEGIN
    DECLARE @monthStart DATE = DATEFROMPARTS(@year, @month, 1);
    DECLARE @monthEnd   DATE = EOMONTH(@monthStart);
    DECLARE @firstDay   DATE = CASE WHEN CAST(@activeFrom AS DATE) > @monthStart
                                    THEN CAST(@activeFrom AS DATE) ELSE @monthStart END;
    DECLARE @lastDay    DATE = CASE WHEN @activeTo IS NOT NULL AND CAST(@activeTo AS DATE) < @monthEnd
                                    THEN CAST(@activeTo AS DATE) ELSE @monthEnd END;
    DECLARE @fraction   DECIMAL(9,6) = 0;

    IF @firstDay <= @lastDay
        SET @fraction = (DATEDIFF(DAY, @firstDay, @lastDay) + 1) * 1.0 / DAY(@monthEnd);

    RETURN @fraction;
END;
GO


-- ---------------------------------------------------------
-- 9b. Function: total monthly bill of a subscriber with a breakdown
--     Returns ONE row: every charge in its own column and the total.
--
--     How the bill is built (assumptions are marked "ASSUMPTION"):
--     1. planFee  = Plan.price x the part of the month the contract was valid.
--        ASSUMPTION: a data plan included in the plan (Plan.internetDataPlanID)
--        is already covered by the plan fee.
--     2. Calls / SMS / MMS made in Lithuania (no foreign place, no foreign
--        recipient): units x PlanBundle.price of the plan.
--        Units: minutes (seconds / 60.0) for calls, 1 per SMS / MMS.
--        BENEFITS: the first Discount.quantity units of the bundle are charged
--        at Discount.price (0.00 = free), scaled by the part of the month the
--        contract was valid.
--        DISCOUNTS: ordered discount packages (OrderedDiscount) add
--        DiscountPackage.quantity free units of that service, but their
--        price is added to the bill. Both quantity and price are scaled by the
--        part of the month the package was valid (half a month = half of the
--        units and half of the price). Plan benefits are used first.
--     3. Calls / SMS / MMS with a foreign place or recipient (roaming and
--        foreign networks): units x ServicePriceInZone.price of the operator,
--        the service and the zone of the foreign country (the initiation
--        place if there is one, otherwise the recipient country); if that price
--        is missing, PlanBundle.price is used. No free units are applied.
--     4. internetCharge = monthly price of ordered internet data plans (scaled
--        by the days they were valid) + data used above the total allowance
--        (plan's data plan + ordered plans, scaled) x additionalPrice per MB.
--        ASSUMPTION: additionalPrice is the price of 1 extra MB.
--     5. additionalServicesCharge = monthly price of ordered additional
--        services, scaled by the days they were valid.
--     A service without any price in the data is charged 0.
--     Subscriber.discount (Assignment 4) is not used.
--     discountSavings is informational (what the free units would have cost);
--     it is NOT part of totalAmount.
-- ---------------------------------------------------------
CREATE OR ALTER FUNCTION dbo.fn_MonthlyBill
(
    @subscriberID INT,
    @year         INT,
    @month        INT
)
RETURNS @bill TABLE
(
    subscriberID             INT,
    billYear                 INT,
    billMonth                INT,
    planFee                  DECIMAL(12,2),
    callsCharge              DECIMAL(12,2),
    smsCharge                DECIMAL(12,2),
    mmsCharge                DECIMAL(12,2),
    internetCharge           DECIMAL(12,2),
    additionalServicesCharge DECIMAL(12,2),
    discountPackagesCharge   DECIMAL(12,2),
    totalAmount              DECIMAL(12,2),
    discountSavings          DECIMAL(12,2)
)
AS
BEGIN
    DECLARE @monthStart  DATE = DATEFROMPARTS(@year, @month, 1);
    DECLARE @nextMonth   DATE = DATEADD(MONTH, 1, @monthStart);

    DECLARE @planID       INT;
    DECLARE @operatorID   INT;
    DECLARE @planPrice    DECIMAL(9,2);
    DECLARE @planDataID   INT;
    DECLARE @planFraction DECIMAL(9,6);

    SELECT @planID       = s.planID,
           @operatorID   = p.operatorID,
           @planPrice    = p.price,
           @planDataID   = p.internetDataPlanID,
           @planFraction = dbo.fn_ActiveMonthFraction(s.contractStartDate, s.contractEndDate, @year, @month)
    FROM dbo.Subscriber s
    JOIN dbo.[Plan] p ON p.planID = s.planID
    WHERE s.subscriberID = @subscriberID;

    IF @planID IS NULL
        RETURN;     -- unknown subscriber: empty result

    -- charges, kept with 2 decimals
    DECLARE @planFee           DECIMAL(12,2) = @planPrice * @planFraction;
    DECLARE @callsCharge       DECIMAL(12,2) = 0;
    DECLARE @smsCharge         DECIMAL(12,2) = 0;
    DECLARE @mmsCharge         DECIMAL(12,2) = 0;
    DECLARE @internetCharge    DECIMAL(12,2) = 0;
    DECLARE @additionalCharge  DECIMAL(12,2) = 0;
    DECLARE @packagesCharge    DECIMAL(12,2) = 0;
    DECLARE @savings           DECIMAL(12,2) = 0;

    -- ordered discount packages valid in the month: free units and price, scaled
    DECLARE @pkg TABLE (serviceID INT, units DECIMAL(14,4), fee DECIMAL(14,4));

    INSERT INTO @pkg (serviceID, units, fee)
    SELECT pb.serviceID,
           dp.quantity * f.fraction,
           dp.price    * f.fraction
    FROM dbo.OrderedDiscount od
    JOIN dbo.DiscountPackage dp ON dp.packageID = od.packageID
    JOIN dbo.PlanBundle pb      ON pb.bundleID = dp.bundleID
    CROSS APPLY (SELECT dbo.fn_ActiveMonthFraction(od.orderDate, od.cancellationDate, @year, @month) AS fraction) f
    WHERE od.subscriberID = @subscriberID
      AND f.fraction > 0;

    SELECT @packagesCharge = ISNULL(SUM(fee), 0) FROM @pkg;

    -- charge per service (calls, SMS, MMS)
    DECLARE @svc TABLE (serviceID INT, charge DECIMAL(14,4), savings DECIMAL(14,4));

    -- (a) usage inside Lithuania: plan price, plan benefits, ordered packages
    INSERT INTO @svc (serviceID, charge, savings)
    SELECT u.serviceID,
           cp.coveredPlan * dd.discountPrice
             + (u.units - cp.coveredPlan - ck.coveredPkg) * ISNULL(b.unitPrice, 0),
           cp.coveredPlan * (ISNULL(b.unitPrice, 0) - dd.discountPrice)
             + ck.coveredPkg * ISNULL(b.unitPrice, 0)
    FROM (
        SELECT c.serviceID,
               SUM(CASE WHEN c.serviceID IN (1, 4, 6)
                        THEN DATEDIFF(SECOND, c.connectionStart, c.connectionEnd) / 60.0
                        ELSE 1.0 END) AS units
        FROM dbo.P2PConnection c
        WHERE c.subscriberID = @subscriberID
          AND c.connectionStart >= @monthStart
          AND c.connectionStart <  @nextMonth
          AND c.connectionInitializationPlace IS NULL
          AND c.recipientCountry IS NULL
        GROUP BY c.serviceID
    ) u
    OUTER APPLY (
        SELECT pb.bundleID, pb.price AS unitPrice
        FROM dbo.PlanBundle pb
        WHERE pb.planID = @planID AND pb.serviceID = u.serviceID
    ) b
    OUTER APPLY (
        SELECT ISNULL(SUM(ds.quantity), 0) * @planFraction AS planUnits,
               ISNULL(MIN(ds.price), 0)                    AS discountPrice
        FROM dbo.Discount ds
        WHERE ds.bundleID = b.bundleID
    ) dd
    OUTER APPLY (
        SELECT ISNULL(SUM(pk.units), 0) AS pkgUnits
        FROM @pkg pk
        WHERE pk.serviceID = u.serviceID
    ) k
    CROSS APPLY (
        SELECT CASE WHEN u.units < dd.planUnits THEN u.units ELSE dd.planUnits END AS coveredPlan
    ) cp
    CROSS APPLY (
        SELECT CASE WHEN u.units - cp.coveredPlan < k.pkgUnits
                    THEN u.units - cp.coveredPlan ELSE k.pkgUnits END AS coveredPkg
    ) ck;

    -- (b) usage with a foreign place or recipient: price of the zone, no free units
    INSERT INTO @svc (serviceID, charge, savings)
    SELECT c.serviceID,
           SUM(x.units * ISNULL(spz.price, ISNULL(pb.price, 0))),
           0
    FROM dbo.P2PConnection c
    LEFT JOIN dbo.Country cn
           ON cn.countryCode = COALESCE(c.connectionInitializationPlace, c.recipientCountry)
    LEFT JOIN dbo.ServicePriceInZone spz
           ON spz.operatorID = @operatorID
          AND spz.serviceID  = c.serviceID
          AND spz.zoneID     = cn.zoneID
    LEFT JOIN dbo.PlanBundle pb
           ON pb.planID = @planID
          AND pb.serviceID = c.serviceID
    CROSS APPLY (
        SELECT CASE WHEN c.serviceID IN (1, 4, 6)
                    THEN DATEDIFF(SECOND, c.connectionStart, c.connectionEnd) / 60.0
                    ELSE 1.0 END AS units
    ) x
    WHERE c.subscriberID = @subscriberID
      AND c.connectionStart >= @monthStart
      AND c.connectionStart <  @nextMonth
      AND (c.connectionInitializationPlace IS NOT NULL OR c.recipientCountry IS NOT NULL)
    GROUP BY c.serviceID;

    SELECT @callsCharge = ISNULL(SUM(CASE WHEN serviceID IN (1, 4, 6) THEN charge END), 0),
           @smsCharge   = ISNULL(SUM(CASE WHEN serviceID IN (2, 5)    THEN charge END), 0),
           @mmsCharge   = ISNULL(SUM(CASE WHEN serviceID = 3          THEN charge END), 0),
           @savings     = ISNULL(SUM(savings), 0)
    FROM @svc;

    -- internet: ordered data plans + data above the allowance
    DECLARE @usedMB          DECIMAL(18,2) = 0;
    DECLARE @allowanceMB     DECIMAL(18,2) = 0;
    DECLARE @extraPricePerMB DECIMAL(9,4);
    DECLARE @orderedPlansFee DECIMAL(14,4) = 0;

    SELECT @usedMB = ISNULL(SUM(ic.amountMB), 0)
    FROM dbo.InternetConnection ic
    WHERE ic.subscriberID = @subscriberID
      AND ic.connectionStart >= @monthStart
      AND ic.connectionStart <  @nextMonth;

    SELECT @allowanceMB     = idp.dataAllowanceMB * @planFraction,
           @extraPricePerMB = idp.additionalPrice
    FROM dbo.InternetDataPlan idp
    WHERE idp.internetDataPlanID = @planDataID;

    SELECT @allowanceMB     = @allowanceMB + ISNULL(SUM(idp.dataAllowanceMB * f.fraction), 0),
           @orderedPlansFee = ISNULL(SUM(idp.price * f.fraction), 0),
           @extraPricePerMB = ISNULL(@extraPricePerMB, MIN(idp.additionalPrice))
    FROM dbo.OrderedInternetDataPlan o
    JOIN dbo.InternetDataPlan idp ON idp.internetDataPlanID = o.internetDataPlanID
    CROSS APPLY (SELECT dbo.fn_ActiveMonthFraction(o.orderDate, o.cancellationDate, @year, @month) AS fraction) f
    WHERE o.subscriberID = @subscriberID;

    SET @internetCharge = @orderedPlansFee
        + CASE WHEN @usedMB > @allowanceMB THEN (@usedMB - @allowanceMB) ELSE 0 END
          * ISNULL(@extraPricePerMB, 0);

    -- additional services (monthly price, scaled by the days they were valid)
    SELECT @additionalCharge = ISNULL(SUM(oas.price * f.fraction), 0)
    FROM dbo.OrderedAdditionalService o
    JOIN dbo.OperatorAdditionalService oas
      ON oas.operatorAdditionalServiceID = o.operatorAdditionalServiceID
    CROSS APPLY (SELECT dbo.fn_ActiveMonthFraction(o.orderDate, o.cancellationDate, @year, @month) AS fraction) f
    WHERE o.subscriberID = @subscriberID;

    INSERT INTO @bill
    VALUES (@subscriberID, @year, @month,
            @planFee, @callsCharge, @smsCharge, @mmsCharge, @internetCharge,
            @additionalCharge, @packagesCharge,
            @planFee + @callsCharge + @smsCharge + @mmsCharge + @internetCharge
                     + @additionalCharge + @packagesCharge,
            @savings);

    RETURN;
END;
GO


-- =========================================================
-- Examples (commented out) - handy when presenting
-- =========================================================
-- SELECT TOP 10 * FROM dbo.vw_SubscriberTotalCallDuration ORDER BY totalCallSeconds DESC;
-- SELECT TOP 10 * FROM dbo.vw_PersonLongestCall ORDER BY durationSeconds DESC;
-- SELECT * FROM dbo.vw_TopDataSubscriberPerOperator;
-- SELECT dbo.fn_LocalCallsAmount(5, 2026, 9) AS localCallsSep2026;
-- SELECT * FROM dbo.fn_MutualContacts(2026, 9);
-- SELECT * FROM dbo.fn_MonthlyBill(5, 2026, 10);
-- SELECT dbo.fn_ActiveMonthFraction('2026-10-16', NULL, 2026, 10) AS secondHalfOfOctober;   -- 0.516129 (16 of 31 days)
--
-- Renewal demo that leaves no trace (rolled back):
-- BEGIN TRANSACTION;
--     DECLARE @id INT = (SELECT TOP 1 subscriberID FROM dbo.Subscriber WHERE contractEndDate IS NOT NULL ORDER BY subscriberID);
--     DECLARE @new INT;
--     EXEC dbo.sp_RenewContract @subscriberID = @id, @newSubscriberID = @new OUTPUT;
--     SELECT @id AS oldContract, @new AS newContract;
-- ROLLBACK TRANSACTION;
