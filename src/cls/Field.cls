VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Field"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

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

