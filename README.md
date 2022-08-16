# MergeIt
MergeIt is a program made in java to read from a text file(.txt) line by line and merge it into the template(.docx) given to it creating a new one with that information.

MergeIt is modular and can be designed to run for multiple docx template files and contain more than the default 6 merge fields.

## How to run
### Default run
MergeIt comes with an ExcutableMergeItFile.jar already made that works for 2 templates and 6 merge fields.
By default the location is set to the C drive in a folder called MergeIt. (C:\MergeIt)
The program looks for a .txt file called "mergefile.txt" loads in the data and then grabs depending upon the line data the template for that line.
#### Example of mergefile.txt:
```
Template|mergefield1|mergefield2|mergefield3|mergefield4|mergefield5|mergefield6
TestTemplateA|John Doe|parachute|jetpack|$100.00||
TestTemplateB|John Doe|jetpack|01/01/1990|$50 ACME Gift Card|five lucky winners|
```
It looks for default for these two docx files to merge the data into: "TestTemplateA extracted.docx" "TestTemplateB extracted.docx".
The program will then create a new docx file that will lead with an "A-" or "B-" depending upon the template used.
The rest of the file will be the first merge field an underscore the date an underscore and the time of the creation of the docx.
#### Examples of created merged docx
A-John Doe_06-03-2022_17-37.docx
A-John Doe_06-03-2022_18-02.docx

## How to edit program
To edit program you will need provided java file (MergeIt.java) POI(https://poi.apache.org/download.html) and import it into your favorite editor. I personally used eclipse.
#### Eclipse setup
For eclipse create a new java project and set the execution environment JRE to JavaSE-1.8 then click next
You will need to go to the Libraries tab and add the following external JARs gotten from https://poi.apache.org/download.html
(You will want to download the Binary Distribution. "poi-bin-5.2.2-20220312.zip")
poi-5.2.2.jar
poi-excelant-5.2.2.jar
poi-javadoc-5.2.2.jar
poi-ooxml-5.2.2.jar
poi-ooxml-full-5.2.2.jar
poi-ooxml-lite-5.2.2.jar
poi-scratchpad-5.2.2.jar

Once you have added the external JARs you can click Finish.
You will then need to right click the java project you just created and hover over "Configure" and click on "Convert to Maven Project".
Change the details if you wish or just click Finish.
After editing the java project you will need to edit the pom.xml file that was created.
Above "<build>" and below "<version>0.0.1-SNAPSHOT</version>" write the following if you are using the same version of poi.
```
    <dependencies>
        <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi-ooxml</artifactId>
            <version>5.2.2</version>
        </dependency>
    </dependencies>
```
After that you are able to import MergeIt.java file and start editing it.

## Editing the program
##### open method and openReadMergeFile method should not be edited.

### Add new merge fields
In the inputData method there are 6 variables called mergefieldX, where X is a number 1-6, you should create a new String variable that should be initialized to "null". (String mergefield7 = null;)
Inside the try you need to add another line that will read in the next token to the variable you just created using nextToken (mergefield7 = st.nextToken();)
After that you will need to go down to the chaining if statements located after the catch of the try.
You will need to add the new merge field variable you created to the calling of the template methods inside the if statements.
Example below using "mergefield7" as the new merge field variable:
```
else if (Template.equals("TemplateC")) {
	TemplateC(TemplateBFile, outputLocationB, mergefield1, mergefield2, mergefield3, mergefield4, mergefield5, mergefield6, mergefield7);
}
```
You will then need to go to the methods using the new merge field and update the method to catch the new merge field variable. Example below:
```
public static void TemplateC(String TemplateCFile, String outputLocationC, String mergefield1, String mergefield2, String mergefield3, String mergefield4, String mergefield5, String mergefield6, String mergefield7) throws IOException {
```
Then go into the try located inside the two for loops and add a nested if statement that checks if the template has the merge field format you are adding and there is a merge field with value in that line of data. Then using replace update the mergefield format with the variable you created. Example below:
```
if (docText.contains("{mergefield7}")) {
	if (!mergefield7.equals(null)) {
        docText = docText.replace("{mergefield7}", mergefield7);
    }
}
```
Repeat adding the merge field variable, creation of the nested if statement and replacement of the merge field format with variable to each template method using the new merge field.

You have completed adding a new merge field.

## Add new template
Inside of main add a two new String variables reflecting the output location of the template and location of the docx fiel for the template.
You will then need to pass the two new variables to open. Example below:
```
/*
 * Example new template with new open to reflect new template.
 */
// Path to TestTemplateC
String TemplateCFile = "c:\\MergeIt\\TemplateC.docx";
      
// Path to save locationC
String outputLocationC = "c:\\MergeIt\\C-";
        
open(textFile, TemplateAFile, TemplateBFile, outputLocationA, outputLocationB, TemplateCFile, outputLocationC);
```

Then inside open method you will need to add the two new variables after "String outputLocationB" and then pass them to inputData method:
```
public static void open(String textFile, String TemplateAFile, String TemplateBFile, String outputLocationA, String outputLocationB, String TemplateCFile, String outputLocationC) {
		BufferedReader br = new BufferedReader(openReadMergeFile(textFile));
		inputData(br, TemplateAFile, TemplateBFile, outputLocationA, outputLocationB, TemplateCFile, outputLocationC);
	}
```

Inside of inputData method you will need to add the new variables after ""String outputLocationB". Example below:
```
public static void inputData(BufferedReader bf, String TemplateAFile, String TemplateBFile, String outputLocationA, String outputLocationB, String TemplateCFile, String outputLocationC) {
```

Then go down to the chaining if statements and add a new else if that checks if the first field of data in the line is equal to your new template name. Example below:
```
if (Template.equals("TestTemplateA")) {
	TemplateA(TemplateAFile, outputLocationA, mergefield1, mergefield2, mergefield3, mergefield4, mergefield5, mergefield6);
}
else if (Template.equals("TestTemplateB")) {
	TemplateB(TemplateBFile, outputLocationB, mergefield1, mergefield2, mergefield3, mergefield4, mergefield5, mergefield6);
}
// New template
else if (Template.equals("TemplateC")) {
	TemplateC(TemplateBFile, outputLocationB, mergefield1, mergefield2, mergefield3, mergefield4, mergefield5, mergefield6, mergefield7);
}
```

To speed up the process copy the commented out method called TemplateC or copy below and rename it after your template(TemplateC).
```
public static void TemplateC(String TemplateCFile, String outputLocationC, String mergefield1, String mergefield2, String mergefield3, String mergefield4, String mergefield5, String mergefield6, String mergefield7) throws IOException {

    try (XWPFDocument doc = new XWPFDocument(Files.newInputStream(Paths.get(TemplateCFile)))) {

        List<XWPFParagraph> xwpfParagraphList = doc.getParagraphs();
        //Iterate over paragraph list and check for the replaceable text in each paragraph
        for (XWPFParagraph xwpfParagraph : xwpfParagraphList) {
            for (XWPFRun xwpfRun : xwpfParagraph.getRuns()) {
                String docText = xwpfRun.getText(0);
                //replacement and setting position
                try {
                    if (docText.contains("{mergefield1}")) {
                        if (!mergefield1.equals(null)) {
                            docText = docText.replace("{mergefield1}", mergefield1);
                        }
                    }
                    if (docText.contains("{mergefield2}")) {
                        if (!mergefield2.equals(null)) {
                            docText = docText.replace("{mergefield2}", mergefield2);
                        }
                    }
                    if (docText.contains("{mergefield3}")) {
                        if (!mergefield3.equals(null)) {
                            docText = docText.replace("{mergefield3}", mergefield3);
                        }
                    }
                    if (docText.contains("{mergefield4}")) {
                        if (!mergefield4.equals(null)) {
                            docText = docText.replace("{mergefield4}", mergefield4);
                        }
                    }
                    if (docText.contains("{mergefield5}")) {
                        if (!mergefield5.equals(null)) {
                            docText = docText.replace("{mergefield5}", mergefield5);
                        }
                    }
                    if (docText.contains("{mergefield6}")) {
                        if (!mergefield6.equals(null)) {
                            docText = docText.replace("{mergefield6}", mergefield6);
                        }
                    }
                    if (docText.contains("{mergefield7}")) {
                        if (!mergefield7.equals(null)) {
                            docText = docText.replace("{mergefield7}", mergefield7);
                        }
                    }
                    xwpfRun.setText(docText, 0);
                } catch (NullPointerException e) {
                    /*
                        * Handle null point exception, happens when it can't locate a merge field.
                        */
                }
            }
        }
        DateTimeFormatter dtf = DateTimeFormatter.ofPattern("MM-dd-uuuu_HH-mm");
        LocalDateTime now = LocalDateTime.now();
        
        outputLocationC = outputLocationC + mergefield1 + "_" + dtf.format(now) + ".docx";
        // save the docs
        try (FileOutputStream out = new FileOutputStream(outputLocationC)) {
            doc.write(out);
        } catch (Exception e) {
            System.out.println("Writting doc file exception.");
            e.printStackTrace();
        }

    } catch (Exception e) {
        System.out.println("TemplateB XWPFDocument exception.");
        e.printStackTrace();
    }
}
```

You will want to make sure the method reflects the variables being passed in to it from inputData and the if statements reflect the merge field variables being passed into the method.

You have compeleted adding a new template.

## How the program works
Inside main there are by default 5 uncommented variables that reference specific location for the files. By default it referneces a folder called "MergeIt" in the root of the C drive.
The variables are then passed to open method which then creates a new BufferedReader variable from the returned value of openReadMergeFile method.
openReadMergeFile creates a new File type variable called inFile from the file in the textfile location.
openReadMergeFile grabs the merge file location set via the textfile variable the method will output an error to console if the file does not exist.
openReadMergeFile will then create a new BufferedReader which reads in the file will output and error if there is an issue in creation of the BufferedReader variable.
openReaMergeFile will then return the BufferedReader variable to open method.

open will then pass off the BufferedReader variable to inputData method along with the rest of the variables passed from main.

##### Start of the loop.
inputData method will try to read the BufferedReader variable line by line if it can't it will output an error to console. Then it goes into a while loop to run through each line in the text file. Then it will try to fill in each mergefield variable with data if it can't it will error and then it passess off the mergefields alond with the specific template location and output location to a specific template depending if the Template variable equals the template name found in the text file.

The template method will then grab the docx template file and create a new XWPFDocument variable and loop through the document paragraph by paragraph.
It gets the text and stores it into a String variable called docText. Then it will check to see if the document conatins the {mergefieldX} and then if the margefield variable is not blank it will repalce the {mergefieldX} with the data in side the variable. It then updates the docText with the new data and after going through each paragraph it then creates a new file with the output location being taken from the defined in main. The default format used is "X-Y_MM-DD-YYYY_H-M.docx" where X is the template type, Y is the first merge field, MM is the current month, DD is the current day, YYYY is the current year, H is the current hour in 24 hour format, and M is the current minute in 24 hour format.

It will then continue back to inputData and loop back through the same process stated above until each line in the merge text file is complete.

## Variable definitions
##### open method and openReadMergeFile method should not be edited.
textfile - Path to mergefile.
```
String textFile = "c:\\MergeIt\\mergefile.txt";
```
TemplateAFile - Path to TestTemplateA.
```
String TemplateAFile = "c:\\MergeIt\\TestTemplateA extracted.docx";
```
outputLocationA - Path to save locationA
```
String outputLocationA = "c:\\MergeIt\\A-";
```
TestTemplateB - Path to TestTemplateB
```
String TemplateBFile = "c:\\MergeIt\\TestTemplateB extracted.docx";
```
outputLocationB - Path to save locationB
```
String outputLocationB = "c:\\MergeIt\\B-";
```
Template - Used to determine which template is being used.
```
String Template = null;
```
mergefieldX - Reflects the merge fields 1-6 replacing X with the corasponding number.
```
String mergefield1 = null;
String mergefield2 = null;
String mergefield3 = null;
String mergefield4 = null;
String mergefield5 = null;
String mergefield6 = null;
```