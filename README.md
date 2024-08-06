
# Libreria Mega 

A reading hub with multiple sections for the user such as: favorites section, read later section, disliked section, a history and book recommendations




## Authors

- [Bryan Shaloat Be Barragan Pulido](https://github.com/Bryan-Shaloat-Be)


## How did I do the project?


The proyec was created focused on the administration and reading of books, a hub where you can have all your books organized for when you want to read one. The proyect use a data base with SQL Server and visual Basic 6.0 to create the user interface. 
## Technical Requirements 

The proyect was build with Visual Basic 6.0 and T-SQL(SQL Server)

- Visual Basic 6.0
- T-SQL (SQL Server)
## Installation the project

IMPORTANT u need search Visual Basic 6.0 and install with some steps, since it is not compatible with windows 8/10/11 and more, its old technology. 

Create an .ini file to conect with database like this form: 
```bash
[database]
provider=SQLOLEDB
data_source= "Server name"
initial_catalog= "Database name"
user_id= "User database name"
password= "Your password"
```




## Proyect images

![Home](https://github.com/Bryan-Shaloat-Be/Libreria-Mega/blob/main/Proyect%20Images/Home.JPG)
![Profile](https://github.com/Bryan-Shaloat-Be/Libreria-Mega/blob/main/Proyect%20Images/profile.JPG)


## Database diagram
![Diagram](https://github.com/Bryan-Shaloat-Be/Libreria-Mega/blob/main/Database/Diagrams/Diagram_Library_Database.JPG)

![Model](https://github.com/Bryan-Shaloat-Be/Libreria-Mega/blob/main/Database/Diagrams/Model_Library_database.JPG)
## Database
A database created with SQL Management Studio and SQL Server was used.

- SQL Server -- version: 16.0.1000.6
- SQL Management Studio 20

Tables

- Users
- Books
- Favorites
- Unfavorites
- Wach_later (read)
- History
## Backend

The backend part was created with simple conection with database and simples queries. All bottons have this queries. 
## Sprint process 

General application

- Creating the app was a challenge in the first steps but over time I managed to handle it. The use of forms was something new for me, so making an application with a lot of style and colors was not easy and I did not achieve a very well designed interface in terms of view and colors. 

Backend

- For the back part I think it was easier to handle, I focused more on it and placed all the crud for the entertainment hub, in the same way I managed to create many validations and error handling.

Database

- Creating the database was not much of a problem (simple) but when using it I made some small changes to the restrictions for the favorites, History, read later and unfavorites operations.
## Sprint Review

| What did I do right? | what I didn't do well| What can i do differently? 
|----|-------------------|------|
| All crud work correctly, alll bottons, functions and querys. The proyect have some validations and error handling| Problems whit design UI/UX. The conection with databse was not implemented correctly or rather efficiently. A little problem with default select items in component ListView1 |Improve ui/ux. improve responsive. Created a global connection to database for all forms in the proyect.


## FeedBack

Any constructive feedback will be welcomed and appreciated. send it to me at the following email bryanbarraganpulido@gmail.com