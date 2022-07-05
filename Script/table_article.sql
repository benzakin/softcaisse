drop TABLE [dbo].[Article];
CREATE TABLE [dbo].[Article] (
    [Article_Id]          INT            IDENTITY (1, 1) NOT NULL,
     [Article_Date]        DATETIME       NULL,
   [Article_Rapport_Num] INT            NULL,
    [Article_Nb]          INT            NULL,
     [Article_Name]        NVARCHAR (100) NULL,
   [Article_Price]       FLOAT (53)     NULL,
    PRIMARY KEY CLUSTERED ([Article_Id] ASC)
);

