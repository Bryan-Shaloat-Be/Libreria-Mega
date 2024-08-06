CREATE DATABASE Books;

USE Books;

CREATE TABLE Users(
	ID_User int IDENTITY(1,1) not null,
	U_Name NVARCHAR(150) not null,
	Mail NVARCHAR(120) not null, 
	U_password NVARCHAR(16) not null,
	Preferences NVARCHAR(20) not null,

		CONSTRAINT PK_User PRIMARY KEY
	(
		ID_User ASC
	),

	CONSTRAINT UQ_Mail UNIQUE
	(
		Mail
	)
)

CREATE TABLE Books(
	ID_Book INT IDENTITY(1,1) not null,
	Title NVARCHAR(50) not null,
	B_Description NVARCHAR(300) not null,
	Category NVARCHAR(20) not null,
	B_Year INT not null,
	Pages INT not null,
	URL_img NVARCHAR(300) not null,

		CONSTRAINT PK_ID_Book PRIMARY KEY
		(
			ID_Book ASC	
		),
		
		CONSTRAINT UQ_Title_Book UNIQUE
		(
			Title
		)
);

CREATE TABLE History(
	ID_History INT IDENTITY(1,1) not null,
	ID_User INT not null,
	ID_Book INT not null,

	CONSTRAINT PK_ID_History PRIMARY KEY
	(
		ID_History ASC
	),

	CONSTRAINT FK_User_History FOREIGN KEY(ID_User)
		REFERENCES Users(ID_User),

	CONSTRAINT FK_Book_History FOREIGN KEY(ID_Book)
		REFERENCES Books(ID_Book)
)

CREATE TABLE Favorites(
	ID_Favorites INT IDENTITY(1,1) not null,
	ID_User INT not null,
	ID_Book INT not null

	CONSTRAINT FK_User_Favorites FOREIGN KEY(ID_User)
		REFERENCES Users(ID_User),

	CONSTRAINT FK_Book_Favorites FOREIGN KEY(ID_Book)
		REFERENCES Books(ID_Book)
)


CREATE TABLE Wach_later(
	ID_WL INT IDENTITY(1,1) not null,
	ID_User INT not null,
	ID_Book INT not null,

	CONSTRAINT FK_User_wach_later FOREIGN KEY(ID_User)
		REFERENCES Users(ID_User),

	CONSTRAINT FK_Book_wach_later FOREIGN KEY(ID_Book)
		REFERENCES Books(ID_Book)
)

CREATE TABLE Unfavorites(
	ID_Unfavorites INT IDENTITY(1,1) not null,
	ID_User INT not null,
	ID_Book INT not null,

	CONSTRAINT FK_User_Unfavorites FOREIGN KEY(ID_User)
		REFERENCES Users(ID_User),

	CONSTRAINT FK_Book_Unfavorites FOREIGN KEY(ID_Book)
		REFERENCES Books(ID_Book)
)

CREATE UNIQUE INDEX IX_Unique_User_Movie
ON Favorites (ID_User, ID_Book)

CREATE UNIQUE INDEX IX_Unique_User_Wach_later
ON Wach_later (ID_User, ID_Book)

CREATE UNIQUE INDEX IX_Unique_User_Unfavorites
ON Unfavorites (ID_User, ID_Book)

CREATE UNIQUE INDEX IX_Unique_User_History
ON History (ID_User, ID_Book)