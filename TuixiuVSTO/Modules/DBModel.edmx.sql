
-- --------------------------------------------------
-- Entity Designer DDL Script for SQL Server 2005, 2008, 2012 and Azure
-- --------------------------------------------------
-- Date Created: 02/13/2019 15:44:57
-- Generated from EDMX file: D:\VS-Projects\TuixiuVSTO\TuixiuVSTO\Modules\DBModel.edmx
-- --------------------------------------------------

SET QUOTED_IDENTIFIER OFF;
GO
USE [jzyy_pay];
GO
IF SCHEMA_ID(N'dbo') IS NULL EXECUTE(N'CREATE SCHEMA [dbo]');
GO

-- --------------------------------------------------
-- Dropping existing FOREIGN KEY constraints
-- --------------------------------------------------


-- --------------------------------------------------
-- Dropping existing tables
-- --------------------------------------------------

IF OBJECT_ID(N'[dbo].[2017sum]', 'U') IS NOT NULL
    DROP TABLE [dbo].[2017sum];
GO
IF OBJECT_ID(N'[dbo].[2018pay]', 'U') IS NOT NULL
    DROP TABLE [dbo].[2018pay];
GO

-- --------------------------------------------------
-- Creating all tables
-- --------------------------------------------------

-- Creating table 'C2017sum'
CREATE TABLE [dbo].[C2017sum] (
    [机构] nvarchar(254)  NULL,
    [姓名] nvarchar(254)  NULL,
    [职工号] nchar(5)  NOT NULL,
    [身份证号码] nvarchar(254)  NULL,
    [工资] decimal(38,4)  NULL,
    [奖金] decimal(38,4)  NULL,
    [奖励性补贴] decimal(19,4)  NULL,
    [机构编号] nvarchar(254)  NULL,
    [年合计收入] decimal(38,4)  NULL
);
GO

-- Creating table 'C2018pay'
CREATE TABLE [dbo].[C2018pay] (
    [职工号] nchar(5)  NOT NULL,
    [姓名] nvarchar(255)  NULL,
    [科室] nvarchar(255)  NULL,
    [身份证号码] nvarchar(255)  NULL,
    [工资应发合计] float  NULL,
    [全年奖金合计] float  NULL,
    [年终奖和补加款] float  NULL,
    [全年收入合计] float  NULL,
    [去年合计收入] float  NULL
);
GO

-- --------------------------------------------------
-- Creating all PRIMARY KEY constraints
-- --------------------------------------------------

-- Creating primary key on [职工号] in table 'C2017sum'
ALTER TABLE [dbo].[C2017sum]
ADD CONSTRAINT [PK_C2017sum]
    PRIMARY KEY CLUSTERED ([职工号] ASC);
GO

-- Creating primary key on [职工号] in table 'C2018pay'
ALTER TABLE [dbo].[C2018pay]
ADD CONSTRAINT [PK_C2018pay]
    PRIMARY KEY CLUSTERED ([职工号] ASC);
GO

-- --------------------------------------------------
-- Creating all FOREIGN KEY constraints
-- --------------------------------------------------

-- --------------------------------------------------
-- Script has ended
-- --------------------------------------------------