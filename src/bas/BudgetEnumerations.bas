Attribute VB_Name = "BudgetEnumerations"
Option Compare Database


Public Enum BOC

    NS = 0

    Payroll = 10

    FTE = 17

    NonSiteTravel = 21

    SiteTravel = 28

    Expenses = 36

    Contracts = 37

    Grants = 41
    
End Enum


Public Enum Field

    NS = 0

    DunsNumber

    ClosedDate

    '-----------------------------------------------------------------------------------
    '------------------              Procurements             --------------------------
    '-----------------------------------------------------------------------------------

    VendorName

    VendorCode

    SecurityOrg

    DocumentControlNumber

    Description
    
    RequestedBy

    '-----------------------------------------------------------------------------------
    '------------------              Requisitions             --------------------------
    '-----------------------------------------------------------------------------------

    RequestDate

    ModifiedBy

    CreatedBy

    DocumentDate

    RequestNumber

    '-----------------------------------------------------------------------------------
    '------------------              Paymentss              ----------------------------
    '-----------------------------------------------------------------------------------

    ContractNumber

    OrderNumber

    CheckDate

    ModificationNumber

    InvoiceDate

    InvoiceNumber

    '-----------------------------------------------------------------------------------
    '------------------          Program Elements           ----------------------------
    '-----------------------------------------------------------------------------------

    Code

    Names

    Title

    '-----------------------------------------------------------------------------------
    '------------------     HumanResourceOrganizations      ----------------------------
    '-----------------------------------------------------------------------------------

    HrOrgCode

    HumanResourceOrganizationCode

    HrOrgName

    HumanResourceOrganizationName

    '-----------------------------------------------------------------------------------
    '------------------          WorkCodes                  ----------------------------
    '-----------------------------------------------------------------------------------

    WorkCode

    WorkCodeName

    ShortName

    ChargeType

    Notifications

    ApproverUserName

    ApprovedDate

    ModifierUserName

    ModifiedDate

    WorkProjectCode

    WorkProjectName

    Percentage

    '-----------------------------------------------------------------------------------
    '------------------      PayrollHours                   ----------------------------
    '-----------------------------------------------------------------------------------

    EndDate

    ApprovalDate

    EmployeeNumber

    EmployeeFirstName

    EmployeeLastName

    Date

    ReportingCode

    ReportingCodeName

    Hours

    '-----------------------------------------------------------------------------------
    '------------------      TravelObligations              ----------------------------
    '-----------------------------------------------------------------------------------

    Destination

    MiddleName

    Address

    DepartureDate

    ReturnDate

    '-----------------------------------------------------------------------------------
    '------------------              Purchases              ----------------------------
    '-----------------------------------------------------------------------------------
    
    DCN
    
    DocumentType

    DocumentPrefix

    OriginalActionDate

    GrantNumber

    ObligatingDocumentNumber

    System

    TransactionNumber

    PurchaseRequest

    '-----------------------------------------------------------------------------------
    '------------------                Employees            ----------------------------
    '-----------------------------------------------------------------------------------

    FirstName

    LastName

    Section

    Email

    Office

    PhoneNumber

    CellNumber

    Status

    '-----------------------------------------------------------------------------------
    '------------------          WorkForceData              ----------------------------
    '-----------------------------------------------------------------------------------

    EmployeeName

    ServiceDate
    
    HireDate

    JobTitle

    OccupationalSeries

    Grade

    Step

    GradeEntryDate

    StepEntryDate

    WigiDueDate

    AppointmentAuthority

    AppointmentType

    BargainingUnit
    
    EmployeeStatus

    RetirementPlan

    '-----------------------------------------------------------------------------------
    '------------------        Reimbursables                ----------------------------
    '-----------------------------------------------------------------------------------

    ReimbursableAgreementNumber

    AgreementNumber

    '-----------------------------------------------------------------------------------
    '------------------                Transfers            ----------------------------
    '-----------------------------------------------------------------------------------

    DocType

    DocumentNumber

    ProcessedDate

    ResourceType

    Lines

    Subline

    FromTo

    Purpose

    '-----------------------------------------------------------------------------------
    '------------------    PayrollObligations               ----------------------------
    '-----------------------------------------------------------------------------------

    PayPeriod

    CalendarDate

    '-----------------------------------------------------------------------------------
    '------------------          Allocations                ----------------------------
    '-----------------------------------------------------------------------------------

    AllocationId

    '-----------------------------------------------------------------------------------
    '------------------        ControlNumbers               ----------------------------
    '-----------------------------------------------------------------------------------

    RegionControlNumber

    FundControlNumber

    BudgetControlNumber

    DivisionControlNumber

    DateIssued

    '-----------------------------------------------------------------------------------
    '------------------         ProgramProjects             ----------------------------
    '-----------------------------------------------------------------------------------

    ProgramProjectCode

    ProgramProjectName

    Laws

    Narrative

    Definition

    '-----------------------------------------------------------------------------------
    '------------------                ProgramResultsCode   ----------------------------
    '-----------------------------------------------------------------------------------

    BudgetLevel

    '-----------------------------------------------------------------------------------
    '------------------              Accounts               ----------------------------
    '-----------------------------------------------------------------------------------

    AccountCode

    AccountName

    '-----------------------------------------------------------------------------------
    '------------------    AppropriationBills               ----------------------------
    '-----------------------------------------------------------------------------------

    PublicLaw

    EnactedDate

    '-----------------------------------------------------------------------------------
    '------------------       FinanceObjectClass            ----------------------------
    '-----------------------------------------------------------------------------------

    FocCode

    FinanceObjectClassCode

    FocName

    FinanceObjectClassName

    '-----------------------------------------------------------------------------------
    '------------------         BudgetObjectClass           ----------------------------
    '-----------------------------------------------------------------------------------

    BocCode

    BudgetObjectClassCode

    BocName

    BudgetObjectClassName

    '-----------------------------------------------------------------------------------
    '------------------       SubAppropriations          ----------------------------
    '-----------------------------------------------------------------------------------
 
    SubAppropriationCode

    SubAppropriationName

    '-----------------------------------------------------------------------------------
    '------------------   ResponsibilityCenter              ----------------------------
    '-----------------------------------------------------------------------------------

    RcCode

    ResponsibilityCenterCode

    RcName

    ResponsibilityCenterName

    '-----------------------------------------------------------------------------------
    '------------------           Appropriations            ----------------------------
    '-----------------------------------------------------------------------------------

    AppropriationCode

    AppropriationName

    '-----------------------------------------------------------------------------------
    '------------------            Funds                    ----------------------------
    '-----------------------------------------------------------------------------------

    FundCode

    FundName

    TreasurySymbol

    '-----------------------------------------------------------------------------------
    '------------------          Activity                   ----------------------------
    '-----------------------------------------------------------------------------------

    ActivityCode

    ActivityName

    '-----------------------------------------------------------------------------------
    '------------------            FiscalYears              ----------------------------
    '-----------------------------------------------------------------------------------

    BFY

    BBFY

    EBFY


    FirstYear

    LastYear

    CancellationDate

    ExpiringYear

    Availability

    FiscalYear

    WorkDays

    WeekDays

    WeekEnds

    '-----------------------------------------------------------------------------------
    '------------------             Holidys                 ----------------------------
    '-----------------------------------------------------------------------------------

    NewYears

    MartinLutherKing

    Presidents

    Memorial

    Independence

    Veterans

    Labor

    Columbus

    Thanksgiving

    Christmas

    '-----------------------------------------------------------------------------------
    '------------------             Supplemental            ----------------------------
    '-----------------------------------------------------------------------------------

    Types

    Time

    '-----------------------------------------------------------------------------------
    '------------------            NationalPrograms         ----------------------------
    '-----------------------------------------------------------------------------------

    NpmCode

    NationalProjgramCode

    NpmName

    NationalProgramName

    '-----------------------------------------------------------------------------------
    '------------------           Organizations             ----------------------------
    '-----------------------------------------------------------------------------------

    OrgCode

    OrganizationCode

    OrgName

    OrganizationName

    CostOrgCode

    CostOrganizationCode

    CostOrgName

    CostOrganizationName

    '-----------------------------------------------------------------------------------
    '------------------    ResourcePlanningOffices          ----------------------------
    '-----------------------------------------------------------------------------------

    RpioCode

    ResourcePlanningOfficeCode

    RpioName

    ResourcePlanningOfficeName

    '-----------------------------------------------------------------------------------
    '------------------        AllowanceHolders             ----------------------------
    '-----------------------------------------------------------------------------------

    AhCode

    AllowanceHolderCode

    AhName

    AllowanceHolderName

    '-----------------------------------------------------------------------------------
    '------------------              Divisions              ----------------------------
    '-----------------------------------------------------------------------------------

    Division

    DivisionName

    Caption

    '-----------------------------------------------------------------------------------
    '------------------                 Objectives        ----------------------------
    '-----------------------------------------------------------------------------------

    ObjectiveCode

    ObjectiveName

    '-----------------------------------------------------------------------------------
    '------------------                 Goals               ----------------------------
    '-----------------------------------------------------------------------------------

    GoalCode

    GoalName

    '-----------------------------------------------------------------------------------
    '------------------             ProgramAreas            ----------------------------
    '-----------------------------------------------------------------------------------

    ProgramAreaCode

    ProgramAreaName

    '-----------------------------------------------------------------------------------
    '------------------          Projects                   ----------------------------
    '-----------------------------------------------------------------------------------

    ProjectCode

    ProjectName

    SiteProjectCode

    SiteProjectName

    '-----------------------------------------------------------------------------------
    '------------------                   Sites             ----------------------------
    '-----------------------------------------------------------------------------------

    SiteName

    EpaSiteId

    City

    District

    County

    StateName

    StateCode

    StreetAddressLine1

    StreetAddressLine2

    ZipCode

    '-----------------------------------------------------------------------------------
    '------------------     INFORMATION TECHNOLOGY            --------------------------
    '-----------------------------------------------------------------------------------

    CostAreaCode

    CostAreaName
    
End Enum



Public Enum FundCode

    NS = 0

    B

    BR

    BR1

    BR2

    BR3

    T

    TC

    TD

    TR

    TR1

    TR2

    TR2A

    TR2B

    TR3

    TS3

    F

    FC

    FD

    FR

    FS3

    H

    HR

    HC

    HD

    E

    E1

    E1C

    E1D

    E1S3

    E2

    E2C

    E2D

    E3

    E3C

    E3D

    E4

    E4C

    E4D

    E5

    E5C

    E5D

    ZL
End Enum




Public Enum RPIO

    NS = 0

    R01

    R02

    R03

    R04

    R05

    R06

    R07

    R08

    R09

    R10

    R11

    R13

    R16

    R17

    R18

    R20

    R26

    R30

    R35

    R39

    R75

    R77

    R92

    R94

    R95

    R98

    R9B

    R9H

    R9P

    R9R

    R9V

    R9Z

    HQ

    RT
End Enum


Public Enum Source
    NS = 0
    AccountingEvents
    Appropriations
    Activity
    AppropriationBills
    ARD
    AllowanceHolders
    Allocations
    Accounts
    Awards
    Balances
    BudgetObjectClass
    ControlNumbers
    CategoricalGrants
    CarryOver
    Commitments
    Contracts
    CWSRF
    DivisionAuthority
    DWSRF
    DivisionExecution
    Divisions
    Deobligations
    DeepWaterHorizon
    Documents
    EJ
    Employees
    Expenses
    External
    ECAD
    EPM
    ExternalTransfers
    Expenditures
    Funds
    FTE
    FiscalYears
    FinanceObjectClass
    FullTimeUtilization
    Grants
    Goals
    HumanResourceData
    HumanResourceOrganizations
    InformationTechnology
    InternalTransfers
    LUST
    LCARD
    LSASD
    LeaveProjections
    LustSupplemental
    MSR
    MSD
    NewObligationalAuthority
    NationalPrograms
    NonSiteTravel
    OilSpill
    ORC
    ORA
    Objectives
    Organizations
    Outlays
    Obligations
    OpenCommitments
    Overtime
    PayrollObligations
    ProgramAreas
    Payroll
    PayrollHours
    Procurements
    Payments
    PurchaseActivity
    PRC
    ProgramProjects
    Programs
    Purchases
    Reimbursables
    Requisitions
    ResourcePlanningOffices
    ResponsibilityCenters
    Reprogrammings
    RegionAuthority
    STAG
    SF6A
    Sites
    Superfund
    SpecialAccounts
    SuperfundSupplemental
    SpecialProjects
    SiteTravel
    SEMD
    Supplemental
    Transfers
    TravelObligations
    Travel
    Utilization
    ULO
    Vendors
    WD
    WSA
    WCF
    WorkforceData
    WorkCodes
    XA
End Enum

