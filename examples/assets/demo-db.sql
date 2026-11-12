-- Demo database used by the Classic ASP broken-vs-fixed examples.
-- Run in SQL Server Management Studio (or sqlcmd) against a local SQL Server / SQL Server Express.
-- All data below is made-up sample data.

IF DB_ID('RescueDemo') IS NULL CREATE DATABASE RescueDemo;
GO
USE RescueDemo;
GO

IF OBJECT_ID('dbo.Orders') IS NOT NULL DROP TABLE dbo.Orders;
IF OBJECT_ID('dbo.Customers') IS NOT NULL DROP TABLE dbo.Customers;
IF OBJECT_ID('dbo.Guestbook') IS NOT NULL DROP TABLE dbo.Guestbook;
IF OBJECT_ID('dbo.CartItems') IS NOT NULL DROP TABLE dbo.CartItems;
GO

CREATE TABLE dbo.Customers (
    CustomerID INT IDENTITY(1,1) PRIMARY KEY,
    Name       NVARCHAR(100) NOT NULL,
    City       NVARCHAR(100) NOT NULL
);

CREATE TABLE dbo.Orders (
    OrderID    INT IDENTITY(1,1) PRIMARY KEY,
    CustomerID INT NOT NULL REFERENCES dbo.Customers(CustomerID),
    OrderDate  DATETIME NOT NULL,
    Total      DECIMAL(12,2) NOT NULL
);
CREATE INDEX IX_Orders_CustomerID ON dbo.Orders(CustomerID);
CREATE INDEX IX_Orders_OrderDate  ON dbo.Orders(OrderDate);

-- Example 03 (Unicode): NVARCHAR stores any language. VARCHAR would turn Hindi into ????.
CREATE TABLE dbo.Guestbook (
    EntryID  INT IDENTITY(1,1) PRIMARY KEY,
    Name     NVARCHAR(200) NOT NULL,
    PostedAt DATETIME NOT NULL DEFAULT GETDATE()
);

-- Example 04 (sessions): cart kept in the database instead of Session().
CREATE TABLE dbo.CartItems (
    CartToken CHAR(36) NOT NULL,
    ProductID INT      NOT NULL,
    Qty       INT      NOT NULL,
    AddedAt   DATETIME NOT NULL DEFAULT GETDATE(),
    CONSTRAINT PK_CartItems PRIMARY KEY (CartToken, ProductID)
);
GO

INSERT INTO dbo.Customers (Name, City) VALUES
 (N'Asha Rao', N'Pune'), (N'Ravi Kumar', N'Delhi'), (N'José Müller', N'Berlin'),
 (N'Meera Iyer', N'Chennai'), (N'Sam Carter', N'London');

INSERT INTO dbo.Orders (CustomerID, OrderDate, Total) VALUES
 (1, '20260304', 1200.00), (1, '20260403', 850.50), (2, '20260315', 99.00),
 (3, '20260402', 4300.00), (4, '20260410', 720.25), (5, '20260412', 15.75);
GO
