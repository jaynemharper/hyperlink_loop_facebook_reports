# hyperlink_loop_facebook_reports

<br>
This code is usefull for the scenerio when you have a datasheet of links which point to Facebook reports that need to all be downloaded on a regular basis. <br>
<br>
1. Creates dataframes of all data per sheet in document containing hyperlinks.<br>
2. Creates a list of hyperlinks from a specified column per dataframe.<br>
3. Logs in to Facebook Business Manager with specified credentials. <br>
4. Loops through hyperlinks (which automatically download reports to /downloads). <br>
5. Defines send_mail function. <br>
6. Loops through identified files per path and attaches to email. <br>
7. Connects to Gmail server and sends email
