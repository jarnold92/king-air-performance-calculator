## King Air Performance Calculator

### Overview
This software is a Winforms Windows desktop application written in VB.NET for various pre-flight performance calculations for the King Air aircraft used at UND.  Things like weight and balance, take-off and cruise calculations can be performed.  The software will also plot calculation results on a graph and you have the ability to print off a TOLD card.  The software does connect to an outside server in order to retrieve values need for the calculations.  There is no installation package or updating mechanism for the application, it simply runs on the desktop as an executable file.

The software was written by a student at UND who had a computer and engineering background with assistance from their instructor to test and validate the calculations.

The instructor who worked with the student thought the software would be very beneficial to use and a great time saving tool for both students and instructors.  The Foundation purchased the software to use and ASN is the keeper of the source code and will provide support. 

### Quick Info
* Application Size: Small
* Development Status: Maintenance/Support
* Current Point of Contact/Product Manager: Rob Clausen - clausen@aero.und.edu or Ryan Wilson - rwilson@aero.und.edu 
* Current Developers: Brent Lahr, Matthew Hillestad, Eddie Carlson
* Original Creator: Mohammad Shakir - Mohammad.shakir@und.edu
* Platform: Windows Desktop
* Main Technologies: Winforms VB.NET

### Other Notes
* To change the server in the project: In lines 9 and 5378, replace the URL with the URL of the new server.
* The default password for the settings of the software is: kingair123

### Dependencies
* [ASN / King Air Server] (http://git1.aero.und.edu/asn/king-air-server/)  used for weight and balance calculation data

### Developer Tools
* Visual Studio 2015

### Deployment Location
* a .exe file is distributed to the users of the software that they copy to their desktop

### Setup/Installation
1. Build Application using Visual Studio
2. Locate .exe file located in bin/Debug or bin/Release folder 
