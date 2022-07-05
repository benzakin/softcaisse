DROP TABLE [dbo].[Rapport];

CREATE TABLE [dbo].[Rapport] (
    [Rapport_Id]        INT           IDENTITY (1, 1) NOT NULL,
    [Rapport_Num]       INT           NOT NULL,
    [Rapport_Date]      DATETIME2 (7) NULL,
    [Rapport_Total]     FLOAT (53)    NULL,
    [Rapport_TVA_10]    FLOAT (53)    NULL,
    [Rapport_TVA_20]    FLOAT (53)    NULL,
    [Rapport_TVA_55]    FLOAT (53)    NULL,
    [Rapport_TVA_TOTAL] FLOAT (53)    NULL,
    [Rapport_Espece]    FLOAT (53)    NULL,
    [Rapport_CB]        FLOAT (53)    NULL,
    [Rapport_TR]        FLOAT (53)    NULL,
    [Rapport_Uber]      FLOAT (53)    NULL,
    [Rapport_Stripe]    FLOAT (53)    NULL,
    [Rapport_Cheque]    FLOAT (53)    NULL,
    [Rapport_Couvert]   INT           NULL,
    PRIMARY KEY CLUSTERED ([Rapport_Id] ASC)
    );