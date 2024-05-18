# Survey-Data-Upload
I started with this project as a teacher who had to manually type the grades of all my students manually to a web page on our learning management system. Years later I saw an opportunity to use it again when I saw myself in front of a huge table asking about attrition data in a Qualtrics survey from one of our survey vendors. Coded in VBA.

Before running the macro, please make sure that:
• You're able to navigate the page's table by using the TAB key on the keyboard
• When you hit tab on the last column of the page's table it goes to the next row
• You already have the right selection on your excel sheet (usually without the headers)

When you run the macro and are ready to move forward, click the OK button. Please, note that you will have 8 seconds to go to the first cell of the website table where you want your data inserted.


If you need assistance importing the BAS file to your Excel file, follow the instructions below:
1. On Excel, press alt + F11 to open the Visual Basic Editor
2. On the Visual Basic editor, right click your file and then click on "import file":

![image](https://github.com/jacksoncaquino/Survey-Data-Upload/assets/61064363/dc2352e3-3062-4d87-b62b-4096050c544f)

3. Choose your BAS file that you downloaded from this repository
4. You'll now have the PutSurveyDataOnForm macro on your list of macros
5. Select your data, go to view, macros, and then select the macro from the list to run.
