﻿CREATE TABLE [AirAndRadiationAuthority]
(
   [AirAndRadiationAuthorityId] INT NOT NULL,
   [PrcId] INT NOT NULL,
   [BudgetLevel] NVARCHAR(255),
   [RPIO] NVARCHAR(255),
   [BFY] NVARCHAR(255),
   [FundCode] NVARCHAR(255),
   [AhCode] NVARCHAR(255),
   [OrgCode] NVARCHAR(255),
   [AccountCode] NVARCHAR(255),
   [RcCode] NVARCHAR(255),
   [BocCode] NVARCHAR(255),
   [Amount] FLOAT,
   [FundName] NVARCHAR(255),
   [BocName] NVARCHAR(255),
   [Division] NVARCHAR(255),
   [DivisionName] NVARCHAR(255),
   [ActivityCode] NVARCHAR(255),
   [NpmName] NVARCHAR(255),
   [NpmCode] NVARCHAR(255),
   [ProgramProjectCode] NVARCHAR(255),
   [ProgramProjectName] NVARCHAR(255),
   [ProgramAreaCode] NVARCHAR(255),
   [ProgramAreaName] NVARCHAR(255),
   [GoalCode] NVARCHAR(255),
   [GoalName] NVARCHAR(255),
   [ObjectiveCode] NVARCHAR(255),
   [ObjectiveName] NVARCHAR(255),
   [ChangeDate] DATETIME
);

CREATE TABLE [AirAndRadiationObligations]
(
   [AirAndRadiationDivisionObligationId] INT NOT NULL,
   [PurchaseId] INT NOT NULL,
   [RpioCode] NVARCHAR(255),
   [BFY] NVARCHAR(255),
   [RcCode] NVARCHAR(255),
   [DivisionName] NVARCHAR(255),
   [AhCode] NVARCHAR(255),
   [AhName] NVARCHAR(255),
   [FundCode] NVARCHAR(255),
   [FundName] NVARCHAR(255),
   [AccountCode] NVARCHAR(255),
   [ActivityCode] NVARCHAR(255),
   [BocCode] NVARCHAR(255),
   [BocName] NVARCHAR(255),
   [NpmCode] NVARCHAR(255),
   [NpmName] NVARCHAR(255),
   [OrgCode] NVARCHAR(255),
   [ProgramProjectCode] NVARCHAR(255),
   [ProgramProjectName] NVARCHAR(255),
   [ProgramAreaCode] NVARCHAR(255),
   [ProgramAreaName] NVARCHAR(255),
   [DocumentControlNumbers] NVARCHAR(255),
   [ReimbursableAgreementNumber] NVARCHAR(255),
   [SiteProjectCode] NVARCHAR(255),
   [DcnPrefix] NVARCHAR(255),
   [DocType] NVARCHAR(255),
   [FocCode] NVARCHAR(255),
   [FocName] NVARCHAR(255),
   [OriginalActionDate] DATETIME,
   [LastActionDate] DATETIME,
   [Commitments] FLOAT NOT NULL,
   [OpenCommitments] FLOAT NOT NULL,
   [Obligations] FLOAT NOT NULL,
   [Deobligations] FLOAT NOT NULL,
   [ULO] FLOAT NOT NULL,
   [Expenditures] FLOAT NOT NULL,
   [Used] FLOAT NOT NULL
);

CREATE TABLE [EnforcementAndComplianceAssuranceAuthority]
(
   [EnforcementAndComplianceAssuranceAuthorityId] INT NOT NULL,
   [PrcId] INT NOT NULL,
   [BudgetLevel] NVARCHAR(255),
   [RPIO] NVARCHAR(255),
   [BFY] NVARCHAR(255),
   [FundCode] NVARCHAR(255),
   [AhCode] NVARCHAR(255),
   [OrgCode] NVARCHAR(255),
   [AccountCode] NVARCHAR(255),
   [RcCode] NVARCHAR(255),
   [BocCode] NVARCHAR(255),
   [Amount] FLOAT,
   [FundName] NVARCHAR(255),
   [BocName] NVARCHAR(255),
   [Division] NVARCHAR(255),
   [DivisionName] NVARCHAR(255),
   [ActivityCode] NVARCHAR(255),
   [NpmName] NVARCHAR(255),
   [NpmCode] NVARCHAR(255),
   [ProgramProjectCode] NVARCHAR(255),
   [ProgramProjectName] NVARCHAR(255),
   [ProgramAreaCode] NVARCHAR(255),
   [ProgramAreaName] NVARCHAR(255),
   [GoalCode] NVARCHAR(255),
   [GoalName] NVARCHAR(255),
   [ObjectiveCode] NVARCHAR(255),
   [ObjectiveName] NVARCHAR(255),
   [ChangeDate] DATETIME
);

CREATE TABLE [EnvironmentalJusticeAuthority]
(
   [EnvironmentalJusticeAuthorityId] INT NOT NULL,
   [PrcId] INT NOT NULL,
   [BudgetLevel] NVARCHAR(255),
   [RPIO] NVARCHAR(255),
   [BFY] NVARCHAR(255),
   [FundCode] NVARCHAR(255),
   [AhCode] NVARCHAR(255),
   [OrgCode] NVARCHAR(255),
   [AccountCode] NVARCHAR(255),
   [RcCode] NVARCHAR(255),
   [BocCode] NVARCHAR(255),
   [Amount] FLOAT,
   [FundName] NVARCHAR(255),
   [BocName] NVARCHAR(255),
   [Division] NVARCHAR(255),
   [DivisionName] NVARCHAR(255),
   [ActivityCode] NVARCHAR(255),
   [NpmName] NVARCHAR(255),
   [NpmCode] NVARCHAR(255),
   [ProgramProjectCode] NVARCHAR(255),
   [ProgramProjectName] NVARCHAR(255),
   [ProgramAreaCode] NVARCHAR(255),
   [ProgramAreaName] NVARCHAR(255),
   [GoalCode] NVARCHAR(255),
   [GoalName] NVARCHAR(255),
   [ObjectiveCode] NVARCHAR(255),
   [ObjectiveName] NVARCHAR(255),
   [ChangeDate] DATETIME
);

CREATE TABLE [EnvironmentalJusticeObligations]
(
   [EnvironmentalJusticeObligationId] UNIQUEIDENTIFIER NOT NULL,
   [PurchaseId] INT NOT NULL,
   [RpioCode] NVARCHAR(255),
   [BFY] NVARCHAR(255),
   [RcCode] NVARCHAR(255),
   [DivisionName] NVARCHAR(255),
   [AhCode] NVARCHAR(255),
   [AhName] NVARCHAR(255),
   [FundCode] NVARCHAR(255),
   [FundName] NVARCHAR(255),
   [AccountCode] NVARCHAR(255),
   [ActivityCode] NVARCHAR(255),
   [BocCode] NVARCHAR(255),
   [BocName] NVARCHAR(255),
   [NpmCode] NVARCHAR(255),
   [NpmName] NVARCHAR(255),
   [OrgCode] NVARCHAR(255),
   [ProgramProjectCode] NVARCHAR(255),
   [ProgramProjectName] NVARCHAR(255),
   [ProgramAreaCode] NVARCHAR(255),
   [ProgramAreaName] NVARCHAR(255),
   [DocumentControlNumbers] NVARCHAR(255),
   [ReimbursableAgreementNumber] NVARCHAR(255),
   [SiteProjectCode] NVARCHAR(255),
   [DcnPrefix] NVARCHAR(255),
   [DocType] NVARCHAR(255),
   [FocCode] NVARCHAR(255),
   [FocName] NVARCHAR(255),
   [OriginalActionDate] DATETIME,
   [LastActionDate] DATETIME,
   [Commitments] FLOAT NOT NULL,
   [OpenCommitments] FLOAT NOT NULL,
   [Obligations] FLOAT NOT NULL,
   [Deobligations] FLOAT NOT NULL,
   [ULO] FLOAT NOT NULL,
   [Expenditures] FLOAT NOT NULL,
   [Used] FLOAT NOT NULL
);

CREATE TABLE [ExternalAffairsAuthority]
(
   [ExternalAffairsAuthorityId] INT NOT NULL,
   [PrcId] INT NOT NULL,
   [BudgetLevel] NVARCHAR(255),
   [RPIO] NVARCHAR(255),
   [BFY] NVARCHAR(255),
   [FundCode] NVARCHAR(255),
   [AhCode] NVARCHAR(255),
   [OrgCode] NVARCHAR(255),
   [AccountCode] NVARCHAR(255),
   [RcCode] NVARCHAR(255),
   [BocCode] NVARCHAR(255),
   [Amount] FLOAT,
   [FundName] NVARCHAR(255),
   [BocName] NVARCHAR(255),
   [Division] NVARCHAR(255),
   [DivisionName] NVARCHAR(255),
   [ActivityCode] NVARCHAR(255),
   [NpmName] NVARCHAR(255),
   [NpmCode] NVARCHAR(255),
   [ProgramProjectCode] NVARCHAR(255),
   [ProgramProjectName] NVARCHAR(255),
   [ProgramAreaCode] NVARCHAR(255),
   [ProgramAreaName] NVARCHAR(255),
   [GoalCode] NVARCHAR(255),
   [GoalName] NVARCHAR(255),
   [ObjectiveCode] NVARCHAR(255),
   [ObjectiveName] NVARCHAR(255),
   [ChangeDate] DATETIME
);

CREATE TABLE [ExternalAffairsObligations]
(
   [ExternalAffairsObligationId] UNIQUEIDENTIFIER NOT NULL,
   [PurchaseId] INT NOT NULL,
   [RpioCode] NVARCHAR(255),
   [BFY] NVARCHAR(255),
   [RcCode] NVARCHAR(255),
   [DivisionName] NVARCHAR(255),
   [AhCode] NVARCHAR(255),
   [AhName] NVARCHAR(255),
   [FundCode] NVARCHAR(255),
   [FundName] NVARCHAR(255),
   [AccountCode] NVARCHAR(255),
   [ActivityCode] NVARCHAR(255),
   [BocCode] NVARCHAR(255),
   [BocName] NVARCHAR(255),
   [NpmCode] NVARCHAR(255),
   [NpmName] NVARCHAR(255),
   [OrgCode] NVARCHAR(255),
   [ProgramProjectCode] NVARCHAR(255),
   [ProgramProjectName] NVARCHAR(255),
   [ProgramAreaCode] NVARCHAR(255),
   [ProgramAreaName] NVARCHAR(255),
   [DocumentControlNumbers] NVARCHAR(255),
   [ReimbursableAgreementNumber] NVARCHAR(255),
   [SiteProjectCode] NVARCHAR(255),
   [DcnPrefix] NVARCHAR(255),
   [DocType] NVARCHAR(255),
   [FocCode] NVARCHAR(255),
   [FocName] NVARCHAR(255),
   [OriginalActionDate] DATETIME,
   [LastActionDate] DATETIME,
   [Commitments] FLOAT NOT NULL,
   [OpenCommitments] FLOAT NOT NULL,
   [Obligations] FLOAT NOT NULL,
   [Deobligations] FLOAT NOT NULL,
   [ULO] FLOAT NOT NULL,
   [Expenditures] FLOAT NOT NULL,
   [Used] FLOAT NOT NULL
);

CREATE TABLE [LaboratoryServicesAndAppliedSciencesObligations]
(
   [LaboratoryServicesAndAppliedSciencesObligationId] INT NOT NULL,
   [PurchaseId] INT NOT NULL,
   [RpioCode] NVARCHAR(255),
   [BFY] NVARCHAR(255),
   [RcCode] NVARCHAR(255),
   [DivisionName] NVARCHAR(255),
   [AhCode] NVARCHAR(255),
   [AhName] NVARCHAR(255),
   [FundCode] NVARCHAR(255),
   [FundName] NVARCHAR(255),
   [AccountCode] NVARCHAR(255),
   [ActivityCode] NVARCHAR(255),
   [BocCode] NVARCHAR(255),
   [BocName] NVARCHAR(255),
   [NpmCode] NVARCHAR(255),
   [NpmName] NVARCHAR(255),
   [OrgCode] NVARCHAR(255),
   [ProgramProjectCode] NVARCHAR(255),
   [ProgramProjectName] NVARCHAR(255),
   [ProgramAreaCode] NVARCHAR(255),
   [ProgramAreaName] NVARCHAR(255),
   [DocumentControlNumbers] NVARCHAR(255),
   [ReimbursableAgreementNumber] NVARCHAR(255),
   [SiteProjectCode] NVARCHAR(255),
   [DcnPrefix] NVARCHAR(255),
   [DocType] NVARCHAR(255),
   [FocCode] NVARCHAR(255),
   [FocName] NVARCHAR(255),
   [OriginalActionDate] DATETIME,
   [LastActionDate] DATETIME,
   [Commitments] FLOAT,
   [OpenCommitments] FLOAT,
   [Obligations] FLOAT,
   [Deobligations] FLOAT,
   [ULO] FLOAT,
   [Expenditures] FLOAT,
   [Used] FLOAT
);

CREATE TABLE [LandChemicalAndRevitalizationAuthority]
(
   [LandChemicalAndRevitalizationAuthorityId] INT NOT NULL,
   [DivisionAuthorityId] INT NOT NULL,
   [BudgetLevel] NVARCHAR(255),
   [RPIO] NVARCHAR(255),
   [BFY] NVARCHAR(255),
   [FundCode] NVARCHAR(255),
   [AhCode] NVARCHAR(255),
   [OrgCode] NVARCHAR(255),
   [AccountCode] NVARCHAR(255),
   [RcCode] NVARCHAR(255),
   [BocCode] NVARCHAR(255),
   [Amount] FLOAT,
   [FundName] NVARCHAR(255),
   [BocName] NVARCHAR(255),
   [Division] NVARCHAR(255),
   [DivisionName] NVARCHAR(255),
   [ActivityCode] NVARCHAR(255),
   [NpmName] NVARCHAR(255),
   [NpmCode] NVARCHAR(255),
   [ProgramProjectCode] NVARCHAR(255),
   [ProgramProjectName] NVARCHAR(255),
   [ProgramAreaCode] NVARCHAR(255),
   [ProgramAreaName] NVARCHAR(255),
   [GoalCode] NVARCHAR(255),
   [GoalName] NVARCHAR(255),
   [ObjectiveCode] NVARCHAR(255),
   [ObjectiveName] NVARCHAR(255),
   [AllocationRatio] FLOAT,
   [ChangeDate] DATETIME
);

CREATE TABLE [LandChemicalAndRevitalizationObligations]
(
   [LandChemicalAndRevitalizationObligationId] INT NOT NULL,
   [PurchaseId] INT NOT NULL,
   [RpioCode] NVARCHAR(255),
   [BFY] NVARCHAR(255),
   [RcCode] NVARCHAR(255),
   [DivisionName] NVARCHAR(255),
   [AhCode] NVARCHAR(255),
   [AhName] NVARCHAR(255),
   [FundCode] NVARCHAR(255),
   [FundName] NVARCHAR(255),
   [AccountCode] NVARCHAR(255),
   [ActivityCode] NVARCHAR(255),
   [BocCode] NVARCHAR(255),
   [BocName] NVARCHAR(255),
   [NpmCode] NVARCHAR(255),
   [NpmName] NVARCHAR(255),
   [OrgCode] NVARCHAR(255),
   [ProgramProjectCode] NVARCHAR(255),
   [ProgramProjectName] NVARCHAR(255),
   [ProgramAreaCode] NVARCHAR(255),
   [ProgramAreaName] NVARCHAR(255),
   [DocumentControlNumbers] NVARCHAR(255),
   [ReimbursableAgreementNumber] NVARCHAR(255),
   [SiteProjectCode] NVARCHAR(255),
   [DcnPrefix] NVARCHAR(255),
   [DocType] NVARCHAR(255),
   [FocCode] NVARCHAR(255),
   [FocName] NVARCHAR(255),
   [OriginalActionDate] DATETIME,
   [Commitments] FLOAT,
   [OpenCommitments] FLOAT,
   [Obligations] FLOAT,
   [Deobligations] FLOAT,
   [ULO] FLOAT,
   [Expenditures] FLOAT,
   [Used] FLOAT
);

CREATE TABLE [MissionSupportAuthority]
(
   [MissionSupportAuthorityId] INT NOT NULL,
   [PrcId] INT NOT NULL,
   [BudgetLevel] NVARCHAR(255),
   [RPIO] NVARCHAR(255),
   [BFY] NVARCHAR(255),
   [FundCode] NVARCHAR(255),
   [AhCode] NVARCHAR(255),
   [OrgCode] NVARCHAR(255),
   [AccountCode] NVARCHAR(255),
   [RcCode] NVARCHAR(255),
   [BocCode] NVARCHAR(255),
   [Amount] FLOAT,
   [FundName] NVARCHAR(255),
   [BocName] NVARCHAR(255),
   [Division] NVARCHAR(255),
   [DivisionName] NVARCHAR(255),
   [ActivityCode] NVARCHAR(255),
   [NpmName] NVARCHAR(255),
   [NpmCode] NVARCHAR(255),
   [ProgramProjectCode] NVARCHAR(255),
   [ProgramProjectName] NVARCHAR(255),
   [ProgramAreaCode] NVARCHAR(255),
   [ProgramAreaName] NVARCHAR(255),
   [GoalCode] NVARCHAR(255),
   [GoalName] NVARCHAR(255),
   [ObjectiveCode] NVARCHAR(255),
   [ObjectiveName] NVARCHAR(255),
   [ChangeDate] DATETIME
);

CREATE TABLE [MissionSupportObligations]
(
   [MissionSupportObligationId] UNIQUEIDENTIFIER NOT NULL,
   [PurchaseId] INT NOT NULL,
   [RpioCode] NVARCHAR(255),
   [BFY] NVARCHAR(255),
   [RcCode] NVARCHAR(255),
   [DivisionName] NVARCHAR(255),
   [AhCode] NVARCHAR(255),
   [AhName] NVARCHAR(255),
   [FundCode] NVARCHAR(255),
   [FundName] NVARCHAR(255),
   [AccountCode] NVARCHAR(255),
   [ActivityCode] NVARCHAR(255),
   [BocCode] NVARCHAR(255),
   [BocName] NVARCHAR(255),
   [NpmCode] NVARCHAR(255),
   [NpmName] NVARCHAR(255),
   [OrgCode] NVARCHAR(255),
   [ProgramProjectCode] NVARCHAR(255),
   [ProgramProjectName] NVARCHAR(255),
   [ProgramAreaCode] NVARCHAR(255),
   [ProgramAreaName] NVARCHAR(255),
   [DocumentControlNumbers] NVARCHAR(255),
   [ReimbursableAgreementNumber] NVARCHAR(255),
   [SiteProjectCode] NVARCHAR(255),
   [DcnPrefix] NVARCHAR(255),
   [DocType] NVARCHAR(255),
   [FocCode] NVARCHAR(255),
   [FocName] NVARCHAR(255),
   [OriginalActionDate] DATETIME,
   [LastActionDate] DATETIME,
   [Commitments] FLOAT NOT NULL,
   [OpenCommitments] FLOAT NOT NULL,
   [Obligations] FLOAT NOT NULL,
   [Deobligations] FLOAT NOT NULL,
   [ULO] FLOAT NOT NULL,
   [Expenditures] FLOAT NOT NULL,
   [Used] FLOAT NOT NULL
);

CREATE TABLE [OfficeOfRegionalCounselAuthority]
(
   [OfficeRegionalCounselAuthorityId] INT NOT NULL,
   [PrcId] INT NOT NULL,
   [BudgetLevel] NVARCHAR(255),
   [RPIO] NVARCHAR(255),
   [BFY] NVARCHAR(255),
   [FundCode] NVARCHAR(255),
   [AhCode] NVARCHAR(255),
   [OrgCode] NVARCHAR(255),
   [AccountCode] NVARCHAR(255),
   [RcCode] NVARCHAR(255),
   [BocCode] NVARCHAR(255),
   [Amount] FLOAT,
   [FundName] NVARCHAR(255),
   [BocName] NVARCHAR(255),
   [Division] NVARCHAR(255),
   [DivisionName] NVARCHAR(255),
   [ActivityCode] NVARCHAR(255),
   [NpmName] NVARCHAR(255),
   [NpmCode] NVARCHAR(255),
   [ProgramProjectCode] NVARCHAR(255),
   [ProgramProjectName] NVARCHAR(255),
   [ProgramAreaCode] NVARCHAR(255),
   [ProgramAreaName] NVARCHAR(255),
   [GoalCode] NVARCHAR(255),
   [GoalName] NVARCHAR(255),
   [ObjectiveCode] NVARCHAR(255),
   [ObjectiveName] NVARCHAR(255),
   [ChangeDate] DATETIME
);

CREATE TABLE [OfficeOfRegionalCounselObligations]
(
   [OfficeOfRegionalCounselObligationId] UNIQUEIDENTIFIER NOT NULL,
   [PurchaseId] INT NOT NULL,
   [RpioCode] NVARCHAR(255),
   [BFY] NVARCHAR(255),
   [RcCode] NVARCHAR(255),
   [DivisionName] NVARCHAR(255),
   [AhCode] NVARCHAR(255),
   [AhName] NVARCHAR(255),
   [FundCode] NVARCHAR(255),
   [FundName] NVARCHAR(255),
   [AccountCode] NVARCHAR(255),
   [ActivityCode] NVARCHAR(255),
   [BocCode] NVARCHAR(255),
   [BocName] NVARCHAR(255),
   [NpmCode] NVARCHAR(255),
   [NpmName] NVARCHAR(255),
   [OrgCode] NVARCHAR(255),
   [ProgramProjectCode] NVARCHAR(255),
   [ProgramProjectName] NVARCHAR(255),
   [ProgramAreaCode] NVARCHAR(255),
   [ProgramAreaName] NVARCHAR(255),
   [DocumentControlNumbers] NVARCHAR(255),
   [ReimbursableAgreementNumber] NVARCHAR(255),
   [SiteProjectCode] NVARCHAR(255),
   [DcnPrefix] NVARCHAR(255),
   [DocType] NVARCHAR(255),
   [FocCode] NVARCHAR(255),
   [FocName] NVARCHAR(255),
   [OriginalActionDate] DATETIME,
   [LastActionDate] DATETIME,
   [Commitments] FLOAT NOT NULL,
   [OpenCommitments] FLOAT NOT NULL,
   [Obligations] FLOAT NOT NULL,
   [Deobligations] FLOAT NOT NULL,
   [ULO] FLOAT NOT NULL,
   [Expenditures] FLOAT NOT NULL,
   [Used] FLOAT NOT NULL
);

CREATE TABLE [OfficeOfTheRegionalAdministratorAuthority]
(
   [OfficeOfTheRegionalAdministratorAuthorityId] INT NOT NULL,
   [PrcId] INT NOT NULL,
   [BudgetLevel] NVARCHAR(255),
   [RPIO] NVARCHAR(255),
   [BFY] NVARCHAR(255),
   [FundCode] NVARCHAR(255),
   [AhCode] NVARCHAR(255),
   [OrgCode] NVARCHAR(255),
   [AccountCode] NVARCHAR(255),
   [RcCode] NVARCHAR(255),
   [BocCode] NVARCHAR(255),
   [Amount] FLOAT,
   [FundName] NVARCHAR(255),
   [BocName] NVARCHAR(255),
   [Division] NVARCHAR(255),
   [DivisionName] NVARCHAR(255),
   [ActivityCode] NVARCHAR(255),
   [NpmName] NVARCHAR(255),
   [NpmCode] NVARCHAR(255),
   [ProgramProjectCode] NVARCHAR(255),
   [ProgramProjectName] NVARCHAR(255),
   [ProgramAreaCode] NVARCHAR(255),
   [ProgramAreaName] NVARCHAR(255),
   [GoalCode] NVARCHAR(255),
   [GoalName] NVARCHAR(255),
   [ObjectiveCode] NVARCHAR(255),
   [ObjectiveName] NVARCHAR(255),
   [ChangeDate] DATETIME
);

CREATE TABLE [OfficeOfTheRegionalAdministratorObligations]
(
   [OfficeOfTheRegionalAdministratorObligationId] UNIQUEIDENTIFIER NOT NULL,
   [PurchaseId] INT NOT NULL,
   [RpioCode] NVARCHAR(255),
   [BFY] NVARCHAR(255),
   [RcCode] NVARCHAR(255),
   [DivisionName] NVARCHAR(255),
   [AhCode] NVARCHAR(255),
   [AhName] NVARCHAR(255),
   [FundCode] NVARCHAR(255),
   [FundName] NVARCHAR(255),
   [AccountCode] NVARCHAR(255),
   [ActivityCode] NVARCHAR(255),
   [BocCode] NVARCHAR(255),
   [BocName] NVARCHAR(255),
   [NpmCode] NVARCHAR(255),
   [NpmName] NVARCHAR(255),
   [OrgCode] NVARCHAR(255),
   [ProgramProjectCode] NVARCHAR(255),
   [ProgramProjectName] NVARCHAR(255),
   [ProgramAreaCode] NVARCHAR(255),
   [ProgramAreaName] NVARCHAR(255),
   [DocumentControlNumbers] NVARCHAR(255),
   [ReimbursableAgreementNumber] NVARCHAR(255),
   [SiteProjectCode] NVARCHAR(255),
   [DcnPrefix] NVARCHAR(255),
   [DocType] NVARCHAR(255),
   [FocCode] NVARCHAR(255),
   [FocName] NVARCHAR(255),
   [OriginalActionDate] DATETIME,
   [LastActionDate] DATETIME,
   [Commitments] FLOAT NOT NULL,
   [OpenCommitments] FLOAT NOT NULL,
   [Obligations] FLOAT NOT NULL,
   [Deobligations] FLOAT NOT NULL,
   [ULO] FLOAT NOT NULL,
   [Expenditures] FLOAT NOT NULL,
   [Used] FLOAT NOT NULL
);

CREATE TABLE [SuperfundAndEmergencyManagementAuthority]
(
   [SuperfundAndEmergencyManagementAuthorityId] INT NOT NULL,
   [PrcId] INT NOT NULL,
   [BudgetLevel] NVARCHAR(255),
   [RPIO] NVARCHAR(255),
   [BFY] NVARCHAR(255),
   [FundCode] NVARCHAR(255),
   [AhCode] NVARCHAR(255),
   [OrgCode] NVARCHAR(255),
   [AccountCode] NVARCHAR(255),
   [RcCode] NVARCHAR(255),
   [BocCode] NVARCHAR(255),
   [Amount] FLOAT,
   [FundName] NVARCHAR(255),
   [BocName] NVARCHAR(255),
   [Division] NVARCHAR(255),
   [DivisionName] NVARCHAR(255),
   [ActivityCode] NVARCHAR(255),
   [NpmName] NVARCHAR(255),
   [NpmCode] NVARCHAR(255),
   [ProgramProjectCode] NVARCHAR(255),
   [ProgramProjectName] NVARCHAR(255),
   [ProgramAreaCode] NVARCHAR(255),
   [ProgramAreaName] NVARCHAR(255),
   [GoalCode] NVARCHAR(255),
   [GoalName] NVARCHAR(255),
   [ObjectiveCode] NVARCHAR(255),
   [ObjectiveName] NVARCHAR(255),
   [ChangeDate] DATETIME
);

CREATE TABLE [SuperfundAndEmergencyManagementObligations]
(
   [SuperfundAndEmergencyManagementObligationId] UNIQUEIDENTIFIER NOT NULL,
   [PurchaseId] INT NOT NULL,
   [RpioCode] NVARCHAR(255),
   [BFY] NVARCHAR(255),
   [RcCode] NVARCHAR(255),
   [DivisionName] NVARCHAR(255),
   [AhCode] NVARCHAR(255),
   [AhName] NVARCHAR(255),
   [FundCode] NVARCHAR(255),
   [FundName] NVARCHAR(255),
   [AccountCode] NVARCHAR(255),
   [ActivityCode] NVARCHAR(255),
   [BocCode] NVARCHAR(255),
   [BocName] NVARCHAR(255),
   [NpmCode] NVARCHAR(255),
   [NpmName] NVARCHAR(255),
   [OrgCode] NVARCHAR(255),
   [ProgramProjectCode] NVARCHAR(255),
   [ProgramProjectName] NVARCHAR(255),
   [ProgramAreaCode] NVARCHAR(255),
   [ProgramAreaName] NVARCHAR(255),
   [DocumentControlNumbers] NVARCHAR(255),
   [ReimbursableAgreementNumber] NVARCHAR(255),
   [SiteProjectCode] NVARCHAR(255),
   [DcnPrefix] NVARCHAR(255),
   [DocType] NVARCHAR(255),
   [FocCode] NVARCHAR(255),
   [FocName] NVARCHAR(255),
   [OriginalActionDate] DATETIME,
   [LastActionDate] DATETIME,
   [Commitments] FLOAT NOT NULL,
   [OpenCommitments] FLOAT NOT NULL,
   [Obligations] FLOAT NOT NULL,
   [Deobligations] FLOAT NOT NULL,
   [ULO] FLOAT NOT NULL,
   [Expenditures] FLOAT NOT NULL,
   [Used] FLOAT NOT NULL
);

CREATE TABLE [WaterDivisionAuthority]
(
   [WaterDivisionAuthorityId] INT NOT NULL,
   [PrcId] INT NOT NULL,
   [BudgetLevel] NVARCHAR(255),
   [RPIO] NVARCHAR(255),
   [BFY] NVARCHAR(255),
   [FundCode] NVARCHAR(255),
   [AhCode] NVARCHAR(255),
   [OrgCode] NVARCHAR(255),
   [AccountCode] NVARCHAR(255),
   [RcCode] NVARCHAR(255),
   [BocCode] NVARCHAR(255),
   [Amount] FLOAT,
   [FundName] NVARCHAR(255),
   [BocName] NVARCHAR(255),
   [Division] NVARCHAR(255),
   [DivisionName] NVARCHAR(255),
   [ActivityCode] NVARCHAR(255),
   [NpmName] NVARCHAR(255),
   [NpmCode] NVARCHAR(255),
   [ProgramProjectCode] NVARCHAR(255),
   [ProgramProjectName] NVARCHAR(255),
   [ProgramAreaCode] NVARCHAR(255),
   [ProgramAreaName] NVARCHAR(255),
   [GoalCode] NVARCHAR(255),
   [GoalName] NVARCHAR(255),
   [ObjectiveCode] NVARCHAR(255),
   [ObjectiveName] NVARCHAR(255),
   [ChangeDate] DATETIME
);

CREATE TABLE [WaterDivisionObligations]
(
   [WaterDivisionObligationId] UNIQUEIDENTIFIER NOT NULL,
   [PurchaseId] INT NOT NULL,
   [RpioCode] NVARCHAR(255),
   [BFY] NVARCHAR(255),
   [RcCode] NVARCHAR(255),
   [DivisionName] NVARCHAR(255),
   [AhCode] NVARCHAR(255),
   [AhName] NVARCHAR(255),
   [FundCode] NVARCHAR(255),
   [FundName] NVARCHAR(255),
   [AccountCode] NVARCHAR(255),
   [ActivityCode] NVARCHAR(255),
   [BocCode] NVARCHAR(255),
   [BocName] NVARCHAR(255),
   [NpmCode] NVARCHAR(255),
   [NpmName] NVARCHAR(255),
   [OrgCode] NVARCHAR(255),
   [ProgramProjectCode] NVARCHAR(255),
   [ProgramProjectName] NVARCHAR(255),
   [ProgramAreaCode] NVARCHAR(255),
   [ProgramAreaName] NVARCHAR(255),
   [DocumentControlNumbers] NVARCHAR(255),
   [ReimbursableAgreementNumber] NVARCHAR(255),
   [SiteProjectCode] NVARCHAR(255),
   [DcnPrefix] NVARCHAR(255),
   [DocType] NVARCHAR(255),
   [FocCode] NVARCHAR(255),
   [FocName] NVARCHAR(255),
   [OriginalActionDate] DATETIME,
   [LastActionDate] DATETIME,
   [Commitments] FLOAT NOT NULL,
   [OpenCommitments] FLOAT NOT NULL,
   [Obligations] FLOAT NOT NULL,
   [Deobligations] FLOAT NOT NULL,
   [ULO] FLOAT NOT NULL,
   [Expenditures] FLOAT NOT NULL,
   [Used] FLOAT NOT NULL
);

CREATE TABLE [WorkforceSupportAccountAuthority]
(
   [WorkforceSupportAccountAuthorityId] INT NOT NULL,
   [PrcId] INT NOT NULL,
   [BudgetLevel] NVARCHAR(255),
   [RPIO] NVARCHAR(255),
   [BFY] NVARCHAR(255),
   [FundCode] NVARCHAR(255),
   [AhCode] NVARCHAR(255),
   [OrgCode] NVARCHAR(255),
   [AccountCode] NVARCHAR(255),
   [RcCode] NVARCHAR(255),
   [BocCode] NVARCHAR(255),
   [Amount] FLOAT,
   [FundName] NVARCHAR(255),
   [BocName] NVARCHAR(255),
   [Division] NVARCHAR(255),
   [DivisionName] NVARCHAR(255),
   [ActivityCode] NVARCHAR(255),
   [NpmName] NVARCHAR(255),
   [NpmCode] NVARCHAR(255),
   [ProgramProjectCode] NVARCHAR(255),
   [ProgramProjectName] NVARCHAR(255),
   [ProgramAreaCode] NVARCHAR(255),
   [ProgramAreaName] NVARCHAR(255),
   [GoalCode] NVARCHAR(255),
   [GoalName] NVARCHAR(255),
   [ObjectiveCode] NVARCHAR(255),
   [ObjectiveName] NVARCHAR(255),
   [ChangeDate] DATETIME
);

CREATE TABLE [WorkforceSupportAccountObligations]
(
   [WorkforceSupportAccountObligationId] UNIQUEIDENTIFIER NOT NULL,
   [PurchaseId] INT NOT NULL,
   [RpioCode] NVARCHAR(255),
   [BFY] NVARCHAR(255),
   [RcCode] NVARCHAR(255),
   [DivisionName] NVARCHAR(255),
   [AhCode] NVARCHAR(255),
   [AhName] NVARCHAR(255),
   [FundCode] NVARCHAR(255),
   [FundName] NVARCHAR(255),
   [AccountCode] NVARCHAR(255),
   [ActivityCode] NVARCHAR(255),
   [BocCode] NVARCHAR(255),
   [BocName] NVARCHAR(255),
   [NpmCode] NVARCHAR(255),
   [NpmName] NVARCHAR(255),
   [OrgCode] NVARCHAR(255),
   [ProgramProjectCode] NVARCHAR(255),
   [ProgramProjectName] NVARCHAR(255),
   [ProgramAreaCode] NVARCHAR(255),
   [ProgramAreaName] NVARCHAR(255),
   [DocumentControlNumbers] NVARCHAR(255),
   [ReimbursableAgreementNumber] NVARCHAR(255),
   [SiteProjectCode] NVARCHAR(255),
   [DcnPrefix] NVARCHAR(255),
   [DocType] NVARCHAR(255),
   [FocCode] NVARCHAR(255),
   [FocName] NVARCHAR(255),
   [OriginalActionDate] DATETIME,
   [LastActionDate] DATETIME,
   [Commitments] FLOAT NOT NULL,
   [OpenCommitments] FLOAT NOT NULL,
   [Obligations] FLOAT NOT NULL,
   [Deobligations] FLOAT NOT NULL,
   [ULO] FLOAT NOT NULL,
   [Expenditures] FLOAT NOT NULL,
   [Used] FLOAT NOT NULL
);

CREATE TABLE [WorkingCapitalFundAuthority]
(
   [WorkingCapitalFundAuthorityId] INT NOT NULL,
   [PrcId] INT NOT NULL,
   [BudgetLevel] NVARCHAR(255),
   [RPIO] NVARCHAR(255),
   [BFY] NVARCHAR(255),
   [FundCode] NVARCHAR(255),
   [AhCode] NVARCHAR(255),
   [OrgCode] NVARCHAR(255),
   [AccountCode] NVARCHAR(255),
   [RcCode] NVARCHAR(255),
   [BocCode] NVARCHAR(255),
   [Amount] FLOAT,
   [FundName] NVARCHAR(255),
   [BocName] NVARCHAR(255),
   [Division] NVARCHAR(255),
   [DivisionName] NVARCHAR(255),
   [ActivityCode] NVARCHAR(255),
   [NpmName] NVARCHAR(255),
   [NpmCode] NVARCHAR(255),
   [ProgramProjectCode] NVARCHAR(255),
   [ProgramProjectName] NVARCHAR(255),
   [ProgramAreaCode] NVARCHAR(255),
   [ProgramAreaName] NVARCHAR(255),
   [GoalCode] NVARCHAR(255),
   [GoalName] NVARCHAR(255),
   [ObjectiveCode] NVARCHAR(255),
   [ObjectiveName] NVARCHAR(255),
   [ChangeDate] DATETIME
);

CREATE TABLE [WorkingCapitalFundObligations]
(
   [WorkingCapitalFundObligationId] UNIQUEIDENTIFIER NOT NULL,
   [PurchaseId] INT NOT NULL,
   [RpioCode] NVARCHAR(255),
   [BFY] NVARCHAR(255),
   [RcCode] NVARCHAR(255),
   [DivisionName] NVARCHAR(255),
   [AhCode] NVARCHAR(255),
   [AhName] NVARCHAR(255),
   [FundCode] NVARCHAR(255),
   [FundName] NVARCHAR(255),
   [AccountCode] NVARCHAR(255),
   [ActivityCode] NVARCHAR(255),
   [BocCode] NVARCHAR(255),
   [BocName] NVARCHAR(255),
   [NpmCode] NVARCHAR(255),
   [NpmName] NVARCHAR(255),
   [OrgCode] NVARCHAR(255),
   [ProgramProjectCode] NVARCHAR(255),
   [ProgramProjectName] NVARCHAR(255),
   [ProgramAreaCode] NVARCHAR(255),
   [ProgramAreaName] NVARCHAR(255),
   [DocumentControlNumbers] NVARCHAR(255),
   [ReimbursableAgreementNumber] NVARCHAR(255),
   [SiteProjectCode] NVARCHAR(255),
   [DcnPrefix] NVARCHAR(255),
   [DocType] NVARCHAR(255),
   [FocCode] NVARCHAR(255),
   [FocName] NVARCHAR(255),
   [OriginalActionDate] DATETIME,
   [LastActionDate] DATETIME,
   [Commitments] FLOAT NOT NULL,
   [OpenCommitments] FLOAT NOT NULL,
   [Obligations] FLOAT NOT NULL,
   [Deobligations] FLOAT NOT NULL,
   [ULO] FLOAT NOT NULL,
   [Expenditures] FLOAT NOT NULL,
   [Used] FLOAT NOT NULL
);
