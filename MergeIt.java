import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.Scanner;
import java.util.StringTokenizer;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;


import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

public class MergeIt {

    public static void main(String[] args) throws IOException {

        // Path to mergefile.
        String textFile = "c:\\MergeIt\\mergefile.txt";
        
        // Path to TestTemplateA
        String TemplateAFile = "c:\\MergeIt\\TestTemplateA extracted.docx";
        
        // Path to save locationA
        String outputLocationA = "c:\\MergeIt\\A-";
        
        // Path to TestTemplateB
        String TemplateBFile = "c:\\MergeIt\\TestTemplateB extracted.docx";
        
        // Path to save locationA
        String outputLocationB = "c:\\MergeIt\\B-";
        
        open(textFile, TemplateAFile, TemplateBFile, outputLocationA, outputLocationB);
        
        /*
         * Example new template with new open to reflect new template.
         */
        // Path to TestTemplateC
//      String TemplateCFile = "c:\\MergeIt\\TemplateC.docx";
      
      // Path to save locationC
//      String outputLocationC = "c:\\MergeIt\\C-";
        
//        open(textFile, TemplateAFile, TemplateBFile, outputLocationA, outputLocationB, TemplateCFile, outputLocationC);
    }
    
    /*
     * Start of TemplateA Merge
     */
    public static void TemplateA(String TemplateAFile, String outputLocationA, String mergefield1, String mergefield2, String mergefield3, String mergefield4, String mergefield5, String mergefield6) throws IOException {

        try (XWPFDocument doc = new XWPFDocument(Files.newInputStream(Paths.get(TemplateAFile)))) {

            List<XWPFParagraph> xwpfParagraphList = doc.getParagraphs();
            //Iterate over paragraph list and check for the replaceable text in each paragraph
            for (XWPFParagraph xwpfParagraph : xwpfParagraphList) {
                for (XWPFRun xwpfRun : xwpfParagraph.getRuns()) {
                    String docText = xwpfRun.getText(0);
//                    System.out.println(docText); // Debug line.
                    //replacement and setting position
                    try {
//                    	System.out.println(docText); // Debug line.
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
//            System.out.println(dtf.format(now)); // Debug line.
            
            outputLocationA = outputLocationA + mergefield1 + "_" + dtf.format(now) + ".docx";
            // Save the doc with the name "A-"+ mergefield1
            try (FileOutputStream out = new FileOutputStream(outputLocationA)) {
                doc.write(out);
            } catch (Exception e) {
            	System.out.println("Writting doc file exception.");
            	e.printStackTrace();
            }

        } catch (Exception e) {
        	System.out.println("TemplateA XWPFDocument exception.");
        	e.printStackTrace();
        }
    }
    /*
     * End of TemplateA Merge
     */
    
    /*
     * Start of TemplateB Merge
     */
    public static void TemplateB(String TemplateBFile, String outputLocationB, String mergefield1, String mergefield2, String mergefield3, String mergefield4, String mergefield5, String mergefield6) throws IOException {

        try (XWPFDocument doc = new XWPFDocument(Files.newInputStream(Paths.get(TemplateBFile)))) {

            List<XWPFParagraph> xwpfParagraphList = doc.getParagraphs();
            //Iterate over paragraph list and check for the replaceable text in each paragraph
            for (XWPFParagraph xwpfParagraph : xwpfParagraphList) {
                for (XWPFRun xwpfRun : xwpfParagraph.getRuns()) {
                    String docText = xwpfRun.getText(0);
//                    System.out.println(docText); // Debug line.
                    //replacement and setting position
                    try {
//                    	System.out.println(docText); // Debug line.
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
//            System.out.println(dtf.format(now)); // Debug line.
            
            outputLocationB = outputLocationB + mergefield1 + "_" + dtf.format(now) + ".docx";
            // save the docs
            try (FileOutputStream out = new FileOutputStream(outputLocationB)) {
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
    /*
     * End of TemplateB Merge
     */
    
    /*
     * Start of example TemplateC Merge. This example has 7 merge fields.
     */
//    public static void TemplateC(String TemplateCFile, String outputLocationC, String mergefield1, String mergefield2, String mergefield3, String mergefield4, String mergefield5, String mergefield6, String mergefield7) throws IOException {
//
//        try (XWPFDocument doc = new XWPFDocument(Files.newInputStream(Paths.get(TemplateCFile)))) {
//
//            List<XWPFParagraph> xwpfParagraphList = doc.getParagraphs();
//            //Iterate over paragraph list and check for the replaceable text in each paragraph
//            for (XWPFParagraph xwpfParagraph : xwpfParagraphList) {
//                for (XWPFRun xwpfRun : xwpfParagraph.getRuns()) {
//                    String docText = xwpfRun.getText(0);
//                    //replacement and setting position
//                    try {
//                    	if (docText.contains("{mergefield1}")) {
//                    		if (!mergefield1.equals(null)) {
//                    			docText = docText.replace("{mergefield1}", mergefield1);
//                    		}
//                    	}
//                    	if (docText.contains("{mergefield2}")) {
//                    		if (!mergefield2.equals(null)) {
//                    			docText = docText.replace("{mergefield2}", mergefield2);
//                    		}
//                    	}
//                    	if (docText.contains("{mergefield3}")) {
//                    		if (!mergefield3.equals(null)) {
//                    			docText = docText.replace("{mergefield3}", mergefield3);
//                    		}
//                    	}
//                    	if (docText.contains("{mergefield4}")) {
//                    		if (!mergefield4.equals(null)) {
//                    			docText = docText.replace("{mergefield4}", mergefield4);
//                    		}
//                    	}
//                    	if (docText.contains("{mergefield5}")) {
//                    		if (!mergefield5.equals(null)) {
//                    			docText = docText.replace("{mergefield5}", mergefield5);
//                    		}
//                    	}
//                    	if (docText.contains("{mergefield6}")) {
//                    		if (!mergefield6.equals(null)) {
//                    			docText = docText.replace("{mergefield6}", mergefield6);
//                    		}
//                    	}
//                    	if (docText.contains("{mergefield7}")) {
//                    		if (!mergefield7.equals(null)) {
//                    			docText = docText.replace("{mergefield7}", mergefield7);
//                    		}
//                    	}
//                    	xwpfRun.setText(docText, 0);
//                    } catch (NullPointerException e) {
//                    	/*
//                    	 * Handle null point exception, happens when it can't locate a merge field.
//                    	 */
//                    }
//                }
//            }
//            DateTimeFormatter dtf = DateTimeFormatter.ofPattern("MM-dd-uuuu_HH-mm");
//            LocalDateTime now = LocalDateTime.now();
//            
//            outputLocationC = outputLocationC + mergefield1 + "_" + dtf.format(now) + ".docx";
//            // save the docs
//            try (FileOutputStream out = new FileOutputStream(outputLocationC)) {
//                doc.write(out);
//            } catch (Exception e) {
//            	System.out.println("Writting doc file exception.");
//            	e.printStackTrace();
//            }
//
//        } catch (Exception e) {
//        	System.out.println("TemplateB XWPFDocument exception.");
//        	e.printStackTrace();
//        }
//    }
    /*
     * End of TemplateC Merge
     */
    
    /*
     * inputData loads in data based off each line in text file. Then sends it to specific method for creating new docx from specific template. 
     * To add new template: after "outputLocationB" you need to do String {templateFileLocationVariableName} for example "String TemplateCFile"
     * Then after that add String {templateLocationVariableName} for example "String outputLocationC".
     * 
     * To add new template: after "outputLocationB" you need to do String {templateFileLocationVariableName} for example "String TemplateCFile"
     * Then after that add String {templateLocationVariableName} for example "String outputLocationC".
     */
	public static void inputData(BufferedReader bf, String TemplateAFile, String TemplateBFile, String outputLocationA, String outputLocationB) {

		Scanner scan = new Scanner(System.in);


		try {
			// Open the file.

			// read in the first line
			String line = "";
			try {
				line = bf.readLine();
			} catch (IOException e) {
				System.out.println("BufferedReader line exception outside while.");
				e.printStackTrace();
			}
			// while there is more data in the file, process it
			while (line != null) { // more lines
//				 System.out.println("Line" + line); // Debug Line
				StringTokenizer st = new StringTokenizer(line, "|");
				// while (st.hasMoreTokens()) { //more items on each line
				
				/*
				 *  Determines which template the line data should go to.
				 *  If adding more merge fields, you need to declare them here and initialize to null.
				 */
				String Template = null;
				String mergefield1 = null;
				String mergefield2 = null;
				String mergefield3 = null;
				String mergefield4 = null;
				String mergefield5 = null;
				String mergefield6 = null;
//				String mergefield7 = null;
				try {
					
					// Determine Template Type.
					Template = st.nextToken();
					
					// Merge Field 1
					mergefield1 = st.nextToken();
					
					// Merge Field 2
					mergefield2 = st.nextToken();
					
					// Merge Field 3
					mergefield3 = st.nextToken();
					
					// Merge Field 4
					mergefield4 = st.nextToken();
					
					// Merge Field 5
					mergefield5 = st.nextToken();
					
					// Merge Field 6
					mergefield6 = st.nextToken();
					
					/*
					 * Add new mergefields here after declaring and initializing above.
					 * Example commented out.
					 */
//					mergefield7 = st.nextToken();
					
				} catch (NoSuchElementException nsee) {
			    /*
			     * This is left blank to handle an exception if there is no more values.
			     */
				}
				
//				System.out.println(Template + " | " + mergefield1 + " | " + mergefield2 + " | " + mergefield3 + " | " + mergefield4 + " | " +  mergefield5 + " | " + mergefield6); // Debug line.
				
				/*
				 * Add another else if statement an example one, TemplateC, is commented out.
				 * You are comparing the first element in the list to the text inside the quotes. Example of what text file should look like below.
				 * Any extra merge fields should also be passed to methods.
				 * 
				 * Example text file:
				 * TemplateC|mergefield1|mergefield2|mergefield3|mergefield4|mergefield5|mergefield6
				 */
				if (Template.equals("TestTemplateA")) {
					TemplateA(TemplateAFile, outputLocationA, mergefield1, mergefield2, mergefield3, mergefield4, mergefield5, mergefield6);
				}
				else if (Template.equals("TestTemplateB")) {
					TemplateB(TemplateBFile, outputLocationB, mergefield1, mergefield2, mergefield3, mergefield4, mergefield5, mergefield6);
				}
//				else if (Template.equals("TemplateC")) {
//					TemplateC(TemplateBFile, outputLocationB, mergefield1, mergefield2, mergefield3, mergefield4, mergefield5, mergefield6, mergefield7);
//				}
				
				

				// read in the next line
				try {
					line = bf.readLine();
				} catch (IOException e) {
					System.out.println("BufferedReader line exception inside while.");
					e.printStackTrace();
				}
			} // end of reading in the data.

		}
		// catch any other type of exception
		catch (Exception e) {
			System.out.println("Other weird things happened");
			e.printStackTrace();
		} finally {
			try {
				bf.close();
			} catch (Exception e) {
				System.out.println("BufferedReader close exception.");
				e.printStackTrace();
			}
		}

	}
    
	/*
	 * open method loads in mergefile from variable in main. Checks to make sure file actually exists and then passes it off to inputData method for parsing of lines in text file and passing to creation of new docx files from templates.
	 * 
	 * To add new template: after "outputLocationB" you need to do String {templateFileLocationVariableName} for example "String TemplateCFile"
     * Then after that add String {templateLocationVariableName} for example "String outputLocationC".
	 */
	public static void open(String textFile, String TemplateAFile, String TemplateBFile, String outputLocationA, String outputLocationB) {
		BufferedReader br = new BufferedReader(openReadMergeFile(textFile));
		inputData(br, TemplateAFile, TemplateBFile, outputLocationA, outputLocationB);
	}
    
	/*
	 * Tries to open file specific from variable in main.
	 */
    public static BufferedReader openReadMergeFile(String textFile) {
    	// create a file instance for the absolute path
    			File inFile = new File(textFile);
    			if (!inFile.exists()) {
    				System.out.println("That file does not exist");
    				System.exit(0);
    			}
    			BufferedReader in = null;
    			try {
    				in = new BufferedReader(new FileReader(inFile));
    			} catch (IOException e) {
    				System.out.println("openReadMergeFile: new BufferedReader(new FileRader) exception.");
    			}
				return in;
    }

}