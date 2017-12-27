This application converts NORSAR GBF seismic catalogues (http://www.norsardata.no/NDC/bulletins/)
to NORDIC (SEISAN) format.
 
The NORSAR GBF catalogue in form of html files has to be downloaded and saved in some local directory
before the application can be run.

Open menu 'File' and select the first html file you want to start from. Then click button 'Select+Convert'
and the application will read all files in that directory one by one, convert to NORSAR format and write
to one 'txt' file.

Seismic events will be transfered to the NORDIC catalogue just in case the coordinates of the event get
inside the area of interest defined by four coordinates presented in the main window.

The application 'Selevent' was written by Andrius Pacea for his personal needs back in distant 2000 then
author had litle knoledge of OOP and therefore the application was written using procedural style of 
programming. The application was slightly modified in 2014 then NORSAR catalogue format had been changed.

The application was written in Visual Basic using Visual Studio v6.0. The whole VB project is uploaded in
this github repository.