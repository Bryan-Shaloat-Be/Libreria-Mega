use Books;

INSERT INTO Books(Title, B_Description,Category,B_Year,Pages,URL_img)
VALUES ('Maze Runer', 'Grupo de jovenes que intentan salir de un laberinto gigantesco que contiene mounstros gigantes, deben intentar salir antes de se comidos','Ciencia ficcion',2019,800,'D:\Usuario Cat\Desktop\PPXP\Practicas_Visual_Basic\Library\maze_runner.jpg');

INSERT INTO Users(U_Name, Mail, U_password, Preferences) 
VALUES ('Catz','catz@gmail.com','1234','Accion');

SELECT * FROM Unfavorites;
SELECT * FROM Books;
SELECT * FROM Favorites;
SELECT * FROM Users;
SELECT * FROM Wach_later;
SELECT * FROM History;

DELETE FROM History
Where ID_History = 16;

use Books;
SELECT 
	Favorites.ID_Favorites,
	Favorites.ID_Book,
	Books.Title,
	Books.B_Description,
	Books.Category,
	Books.Pages,
	Books.B_Year,
	Books.URL_img
from Favorites
join Books on Favorites.ID_Book = Books.ID_Book
where Favorites.ID_User = 1;

SELECT 
	Users.ID_User,
	Users.Preferences,
	Books.Title,
	Books.B_Description,
	Books.Category,
	Books.Pages,
	Books.B_Year,
	Books.URL_img
FROM Users
join Books on Users.Preferences = Books.Category
WHERE Users.ID_User = 1;