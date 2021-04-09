import java.io.*;
import java.util.regex.*;
import com.syncfusion.docio.*;
import com.syncfusion.javahelper.system.text.regularExpressions.MatchSupport;

public class ReplaceTextWithImage {

	public static void main(String[] args) throws Exception {
		//Opens the input Word document.
		WordDocument document = new WordDocument(getDataDir("FindAndReplaceImage_Template.docx"));
		//Finds all the image placeholder text in the Word document.
		TextSelection[] textSelections = document.findAll(Pattern.compile(MatchSupport.trimPattern("^//(.*)")));
		for (int i = 0; i < textSelections.length; i++) {
			// Replaces the image placeholder text with desired image.
			WParagraph paragraph = new WParagraph(document);
			WPicture picture = (WPicture) paragraph
					.appendPicture(new FileInputStream(getDataDir(textSelections[i].getSelectedText() + ".png")));
			TextSelection newSelection = new TextSelection(paragraph, 0, 1);
			TextBodyPart bodyPart = new TextBodyPart(document);
			bodyPart.getBodyItems().add(paragraph);
			document.replace(textSelections[i].getSelectedText(), bodyPart, true, true);
		}
		//Saves the resultant file in the given path.
		document.save("Result.docx");
		document.close();

	}

	/**
	 * Get the file path
	 * 
	 * @param path specifies the file path
	 */
	private static String getDataDir(String path) {
		File dir = new File(System.getProperty("user.dir"));
		if (!(dir.toString().endsWith("Java-Word-Find-and-Replace-Examples")))
			dir = dir.getParentFile();
		dir = new File(dir, "resources");
		dir = new File(dir, path);
		if (dir.isDirectory() == false)
			dir.mkdir();
		return dir.toString();
	}
}
