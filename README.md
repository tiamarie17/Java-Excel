# Java-Excel

The goal of this project was to create a program that takes in an excel spreadsheet, does text analytics on the sheet by analyzing word frequency, and creates and prints a new Excel workbook 
with the text analytics data on it for the user to see. 

To solve this problem, I started by reading documentation about using Java with Excel files. I completed a tutorial that allowed me to create methods for opening and reading 
an Excel file. I modified the method so that it would read all the sheets in a file by using an Iterator. Once I had the data from the sheet, I also needed to access the words from the .txt file
that will be searched for in the document. I created a static method in the Utils folder that converts the .txt file into an array of Strings that can be passed into another method.

Once I had all the data, I created a method called CountWordFrequency to do the text analytics on the documents. It takes in two parameters--the array of Strings (the words we are searching for), and the data from
the Excel spreadsheets as a Map<Integer, List<String>>. I started by converting the array of strings into a hashmap and set all of the values to zero. In this way, I used a hashmap to keep track of the count of each word (key) found in the Excel document. Regarding the data from the Excel spreadsheet, I only needed to access the values within the Map. Since the values are a List<String> type, I was able
to loop through those. I ended up needing to convert an array of objects into a string array. Within the for loops, I checked to see if the words listed in the .txt file existed in the Excel document. If I found a match, I wrote an if statement that added 1 to the count value in the hashmap for that particular word. 

Finally, I added a CreateWorkBook Method that creates a new Excel workbook and sheet with the data returned from the CountWordFrequencyMethod. This method styles the cells and labels the columns.
To set the cell values and get the appropriate number of rows, I looped through the hashmap using the Map.entrySet() method and set the cell values to each key/value pair in the set.
The method also sets the file location, name and type, and closes the workbook once it has been written. 



##Technologies used
Java, Maven

### Next steps
This project really helped me gain a deeper understanding of Collections in Java, specifically the Maps, Lists, and Sets. My next steps for this project are to practice
replacing some of the for loops with Java Streams which I have also been learning about. I also plan to create another version of this project that runs with Spring Boot and includes
REST service so that I can test out the routes in Postman. Finally, I plan to go back and refactor this code as well.
