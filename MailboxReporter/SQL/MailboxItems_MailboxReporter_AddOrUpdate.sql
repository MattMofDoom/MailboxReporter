/* Set this to the DB that will be used */
USE [MailboxReporter]
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[MailboxItems](
	[EmailId] [bigint] IDENTITY(1,1) NOT NULL,
	[Id] [varchar](512) NOT NULL,
	[MailboxAddress] [varchar](512) NOT NULL,
	[InternetMessageId] [varchar](512) NOT NULL,
	[ConversationId] [varchar](512) NOT NULL,
	[SentDate] [datetime] NOT NULL,
	[ReceivedDate] [datetime] NOT NULL,
	[CreatedDate] [datetime] NOT NULL,
	[ModifiedDate] [datetime] NOT NULL,
	[ModifiedName] [nvarchar](256) NOT NULL,	
	[FromName] [nvarchar](256) NOT NULL,
	[FromAddress] [nvarchar](256) NOT NULL,
	[ReplyToName] [nvarchar](max) NULL,
	[ReplyToAddress] [nvarchar](max) NULL,
	[ToName] [nvarchar](max) NULL,
	[ToAddress] [nvarchar](max) NULL,
	[CcName] [nvarchar](max) NULL,
	[CcAddress] [nvarchar](max) NULL,
	[Priority] [nvarchar](10) NULL,
	[Subject] [nvarchar](256) NOT NULL,
	[Body] [nvarchar](max) NULL,
	[BodyType] [nvarchar](10) NOT NULL,
	[Size] [int] NOT NULL,
	[AttachmentCount] [int] NOT NULL,
	[Attachments] [nvarchar](max) NOT NULL,
	[IsRead] [bit] NOT NULL,
 CONSTRAINT [PK_MailboxItems] PRIMARY KEY CLUSTERED 
(
	[EmailId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE NONCLUSTERED INDEX [IX_MailboxData_Address] ON [dbo].[MailboxItems]
(
	[MailboxAddress] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO

CREATE NONCLUSTERED INDEX [IX_MailboxData_ConversationId] ON [dbo].[MailboxItems]
(
	[ConversationId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO

CREATE NONCLUSTERED INDEX [IX_MailboxData_Created] ON [dbo].[MailboxItems]
(
	[CreatedDate] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO

CREATE NONCLUSTERED INDEX [IX_MailboxData_FromAddress] ON [dbo].[MailboxItems]
(
	[FromAddress] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO

CREATE NONCLUSTERED INDEX [IX_MailboxData_Id] ON [dbo].[MailboxItems]
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO

CREATE NONCLUSTERED INDEX [IX_MailboxData_InternetMessageId] ON [dbo].[MailboxItems]
(
	[InternetMessageId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO


CREATE NONCLUSTERED INDEX [IX_MailboxData_IsRead] ON [dbo].[MailboxItems]
(
	[IsRead] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO

CREATE NONCLUSTERED INDEX [IX_MailboxData_Modified] ON [dbo].[MailboxItems]
(
	[ModifiedDate] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO

CREATE NONCLUSTERED INDEX [IX_MailboxData_ModifiedName] ON [dbo].[MailboxItems]
(
	[ModifiedName] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO
CREATE NONCLUSTERED INDEX [IX_MailboxData_Received] ON [dbo].[MailboxItems]
(
	[ReceivedDate] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO

CREATE NONCLUSTERED INDEX [IX_MailboxData_Sent] ON [dbo].[MailboxItems]
(
	[SentDate] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO

CREATE NONCLUSTERED INDEX [IX_MailboxData_Size] ON [dbo].[MailboxItems]
(
	[Size] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO

CREATE PROCEDURE [dbo].[MailboxReporter_AddOrUpdate]
(
	@Id varchar(512),
	@MailboxAddress varchar(512),
	@InternetMessageId varchar(512),
	@ConversationId varchar(512),
	@SentDate datetime,
	@ReceivedDate datetime,
	@CreatedDate datetime,
	@ModifiedDate datetime,
	@ModifiedName nvarchar(256),
	@FromName nvarchar(256),
	@FromAddress nvarchar(256),
	@ReplyToName nvarchar(max),
	@ReplyToAddress nvarchar(max),
	@ToName nvarchar(max),
	@ToAddress nvarchar(max),
	@CcName nvarchar(max),
	@CcAddress nvarchar(max),
	@Priority nvarchar(10),
	@Subject nvarchar(256),
	@Body nvarchar(max),
	@BodyType nvarchar(10),
	@Size int,
	@AttachmentCount int,
	@Attachments nvarchar(max),
	@IsRead bit
)
AS
BEGIN
MERGE [dbo].[MailboxItems] WITH (SERIALIZABLE) AS T
USING (VALUES (@Id, @MailboxAddress, @InternetMessageId, @ConversationId, @SentDate, @ReceivedDate, @CreatedDate, @ModifiedDate, @ModifiedName, @FromName, @FromAddress, @ReplyToName, @ReplyToAddress, @ToName, @ToAddress, @CcName, @CcAddress, @Priority, @Subject, @Body, @BodyType, @Size, @AttachmentCount, @Attachments, @IsRead)) 
AS U (Id, MailboxAddress, InternetMessageId, ConversationId, SentDate, ReceivedDate, CreatedDate, ModifiedDate, ModifiedName, FromName, FromAddress, ReplyToName, ReplyToAddress, ToName, ToAddress, CcName, CcAddress, Priority, Subject, Body, BodyType, Size, AttachmentCount, Attachments, IsRead)
ON U.InternetMessageId = T.InternetMessageId
WHEN MATCHED AND (T.IsRead != U.IsRead OR T.ModifiedDate != U.ModifiedDate OR T.ModifiedName != U.ModifiedName) THEN
 UPDATE SET T.IsRead = U.IsRead, T.ModifiedDate = U.ModifiedDate, T.ModifiedName = U.ModifiedName
WHEN NOT MATCHED THEN
 INSERT(Id, MailboxAddress, InternetMessageId, ConversationId, SentDate, ReceivedDate, CreatedDate, ModifiedDate, ModifiedName, FromName, FromAddress, ReplyToName, ReplyToAddress, ToName, ToAddress, CcName, CcAddress, Priority, Subject, Body, BodyType, Size, AttachmentCount, Attachments, IsRead) 
 VALUES (U.Id, U.MailboxAddress, U.InternetMessageId, U.ConversationId, U.SentDate, U.ReceivedDate, U.CreatedDate, U.ModifiedDate, U.ModifiedName, U.FromName, U.FromAddress, U.ReplyToName, U.ReplyToAddress, U.ToName, U.ToAddress, U.CcName, U.CcAddress, U.Priority, U.Subject, U.Body, U.BodyType, U.Size, U.AttachmentCount, U.Attachments, U.IsRead);
END
GO

/* Need to grant execute permission to the Windows account of the machine running MailboxReporter */
CREATE LOGIN [NT AUTHORITY\SYSTEM] FROM WINDOWS
GO
CREATE USER [NT AUTHORITY\SYSTEM] FOR  LOGIN [NT AUTHORITY\SYSTEM]
GO
GRANT Execute ON dbo.MailboxReporter_AddOrUpdate to [NT AUTHORITY\SYSTEM]
GO
