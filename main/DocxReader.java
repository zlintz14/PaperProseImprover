package main;

import java.awt.Dimension;


import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

import java.awt.Color;
import java.awt.Desktop;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Arrays;
import java.util.List;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextField;

import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRelation;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STUnderline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTColor;

public class DocxReader {

	public static void main(String[] args) {

		//create frame and prompt for user input 
		JFrame main_frame = new JFrame();
		JPanel main_panel = new JPanel();
		JTextField user_word_highlight_num = new JTextField(10);
		JButton ok_btn = new JButton("OK");
		JLabel prompt = new JLabel("A word should be highlighted if it appears this many times in "
				+ "the paper: ");
		main_panel.setPreferredSize(new Dimension(650, 75));
		main_frame.setSize(500, 300);
		main_frame.setLocation(500, 300);
		main_panel.add(prompt, JLabel.LEFT_ALIGNMENT);
		main_panel.add(user_word_highlight_num);
		ok_btn.addActionListener(new ActionListener(){
			@Override
			public void actionPerformed(ActionEvent a){
				if(Integer.parseInt(user_word_highlight_num.getText()) < 1){
					JOptionPane.showMessageDialog(null, "Enter a number greater than 0", "Number Error", JOptionPane.ERROR_MESSAGE);
				}
				else{
					main_frame.dispose();
					//make window for user to choose file
					JFileChooser window = new JFileChooser();
					
					//store returned value of selected button on window 
					int returned_value = window.showOpenDialog(null);
					
					//see if file was selected
					try{
						if(returned_value == JFileChooser.APPROVE_OPTION){
							readDocxFile(window.getSelectedFile(), 
									Integer.parseInt(user_word_highlight_num.getText()));
						}
					}catch(Exception e){
						e.printStackTrace();
					}
				}	
			}
		});
		main_panel.add(ok_btn);
		main_frame.add(main_panel);
		main_frame.pack();
		main_frame.setVisible(true);

	}
	
	public static void readDocxFile(File file, int user_word_highlight_num) {

		try{
			
			//read in selected File
			FileInputStream fis = new FileInputStream(file);
			XWPFDocument document = new XWPFDocument(fis);
			
			List<XWPFParagraph> paragraphs = document.getParagraphs();
			
			//create string array to hold words in each paragraph at each position in array
			//(each position = one paragraph of words)
			String[] doc_words = new String[paragraphs.size()];
			
			for(int i = 0; i < paragraphs.size(); i++){
				doc_words[i] = "";
				doc_words[i] += paragraphs.get(i).getText();
			}
			
			//clone doc_words to new array for doing comparisons so as to not alter original array
			String[] doc_words_copy = doc_words.clone();
			
			//remove all punctuation marks for comparisons 
			for(int i = 0; i < doc_words_copy.length; i++){
				doc_words_copy[i] = removePunctuation(doc_words_copy[i]);
			}
			
			fis.close();
			
			DocxReaderWordHighlighter word_highlighter = new DocxReaderWordHighlighter(doc_words_copy,
					 user_word_highlight_num);
			word_highlighter.putWordsInMap();
			Map<String, Color> highlight_colors = word_highlighter.getHighlightColors();
			
			
			Map<String, String> hex_colors = determineWordHighlight(doc_words, highlight_colors);
			
			//create new doc for edits 
			XWPFDocument new_document = new XWPFDocument();
			
			//create arrays for creating output
			XWPFParagraph[] doc_paragraphs = new XWPFParagraph[paragraphs.size()];
			XWPFRun[] paragraph_runs = new XWPFRun[paragraphs.size()];
			
			for(int i = 0; i < paragraphs.size(); i++){
				doc_paragraphs[i] = new_document.createParagraph();
		    	paragraph_runs[i] = doc_paragraphs[i].createRun();	
		    	//ensures that run is paragraph to be indented and not a header
		    	if(doc_words[i].length() > 150){
		    		paragraph_runs[i].addTab();
		    	}
		    	//this case is most likely title, so underline
		    	if(doc_words[i].length() < 150 && doc_words[i].length() > 30){
		    		paragraph_runs[i].setUnderline(UnderlinePatterns.SINGLE);
		    	}
		 
		    	String[] temp_array = doc_words[i].split(" ");
				for(int x = 0; x < temp_array.length; x++){
					paragraph_runs[i] = doc_paragraphs[i].createRun();
					paragraph_runs[i].setText(" " + temp_array[x]);
					if(hex_colors.containsKey(temp_array[x])){
						CTShd cTShd = paragraph_runs[i].getCTR().addNewRPr().addNewShd();
						cTShd.setVal(STShd.CLEAR);
						cTShd.setColor("auto");
						cTShd.setFill(hex_colors.get(temp_array[x]));
						String temp_word = temp_array[x];
						temp_array[x] = removePunctuation(temp_array[x]);
						temp_array[x] = makeLowerCase(temp_array[x]);
						appendExternalHyperlink("http://www.thesaurus.com/browse/" + temp_array[x] + "?s=t", 
								temp_word, doc_paragraphs[i]);
					}
					paragraph_runs[i].setText(" ");
				}
		    	
			}
			
			//write new Word Docx
		    FileOutputStream fos = new FileOutputStream("Edited docx.docx");	
		    new_document.write(fos);
			
		    fos.close();
			
		    Desktop.getDesktop().open(new File("Edited docx.docx"));
			
			
		}catch(Exception e){
			e.printStackTrace();
		}
		
	}
	
	public static String removePunctuation(String paragraph){
		boolean stillPunctuation = true;
		while(stillPunctuation){
			stillPunctuation = false;
			if(paragraph.indexOf(".") != -1){
				stillPunctuation = true;
				paragraph = paragraph.substring(0, paragraph.indexOf(".")) +
						paragraph.substring(paragraph.indexOf(".") + 1);
			}	
			else if(paragraph.indexOf(",") != -1){
				stillPunctuation = true;
				paragraph = paragraph.substring(0, paragraph.indexOf(",")) +
						paragraph.substring(paragraph.indexOf(",") + 1);
			}	
			else if(paragraph.indexOf("?") != -1){
				stillPunctuation = true;
				paragraph = paragraph.substring(0, paragraph.indexOf("?")) +
						paragraph.substring(paragraph.indexOf("?") + 1);
			}	
			else if(paragraph.indexOf(";") != -1){
				stillPunctuation = true;
				paragraph = paragraph.substring(0, paragraph.indexOf(";")) +
						paragraph.substring(paragraph.indexOf(";") + 1);
			}
			else if(paragraph.indexOf(":") != -1){
				stillPunctuation = true;
				paragraph = paragraph.substring(0, paragraph.indexOf(":")) +
						paragraph.substring(paragraph.indexOf(":") + 1);
			}
			else if(paragraph.indexOf("$") != -1){
				stillPunctuation = true;
				paragraph = paragraph.substring(0, paragraph.indexOf("$")) +
						paragraph.substring(paragraph.indexOf("$") + 1);
			}
			else if(paragraph.indexOf("(") != -1){
				stillPunctuation = true;
				paragraph = paragraph.substring(0, paragraph.indexOf("(")) +
						paragraph.substring(paragraph.indexOf("(") + 1);
			}
			else if(paragraph.indexOf(")") != -1){
				stillPunctuation = true;
				paragraph = paragraph.substring(0, paragraph.indexOf(")")) +
						paragraph.substring(paragraph.indexOf(")") + 1);
			}	
			else if(paragraph.indexOf("[") != -1){
				stillPunctuation = true;
				paragraph = paragraph.substring(0, paragraph.indexOf("[")) +
						paragraph.substring(paragraph.indexOf("[") + 1);
			}
			else if(paragraph.indexOf("]") != -1){
				stillPunctuation = true;
				paragraph = paragraph.substring(0, paragraph.indexOf("]")) +
						paragraph.substring(paragraph.indexOf("]") + 1);
			}
			else if(paragraph.indexOf("*") != -1){
				stillPunctuation = true;
				paragraph = paragraph.substring(0, paragraph.indexOf("*")) +
						paragraph.substring(paragraph.indexOf("*") + 1);
			}
			else if(paragraph.indexOf("\"") != -1){
				stillPunctuation = true;
				paragraph = paragraph.substring(0, paragraph.indexOf("“")) +
						paragraph.substring(paragraph.indexOf("“") + 1);
			}
			else if(paragraph.indexOf("“") != -1){
				stillPunctuation = true;
				paragraph = paragraph.substring(0, paragraph.indexOf("“")) +
						paragraph.substring(paragraph.indexOf("“") + 1);
			}
			else if(paragraph.indexOf("”") != -1){
				stillPunctuation = true;
				paragraph = paragraph.substring(0, paragraph.indexOf("”")) +
						paragraph.substring(paragraph.indexOf("”") + 1);
			}
			else if(paragraph.indexOf("‘") != -1){
				stillPunctuation = true;
				paragraph = paragraph.substring(0, paragraph.indexOf("‘")) +
						paragraph.substring(paragraph.indexOf("‘") + 1);
			}
			else if(paragraph.indexOf("’") != -1){
				stillPunctuation = true;
				paragraph = paragraph.substring(0, paragraph.indexOf("’")) +
						paragraph.substring(paragraph.indexOf("’") + 1);
			}
			else if(paragraph.indexOf("-") != -1){
				stillPunctuation = true;
				paragraph = paragraph.substring(0, paragraph.indexOf("-")) +
						paragraph.substring(paragraph.indexOf("-") + 1);
			}
			else if(paragraph.indexOf("!") != -1){
				stillPunctuation = true;
				paragraph = paragraph.substring(0, paragraph.indexOf("!")) +
						paragraph.substring(paragraph.indexOf("!") + 1);
			}
			else if(paragraph.indexOf("&") != -1){
				stillPunctuation = true;
				paragraph = paragraph.substring(0, paragraph.indexOf("&")) +
						paragraph.substring(paragraph.indexOf("&") + 1);
			}	
			else if(paragraph.indexOf("/") != -1){
				stillPunctuation = true;
				paragraph = paragraph.substring(0, paragraph.indexOf("/")) +
						paragraph.substring(paragraph.indexOf("/") + 1);
			}
			else if(paragraph.indexOf("\\") != -1){
				stillPunctuation = true;
				paragraph = paragraph.substring(0, paragraph.indexOf("\\")) +
						paragraph.substring(paragraph.indexOf("\\") + 1);
			}	
		}	
			
		return paragraph;
	}
	
	public static Map<String, String> determineWordHighlight(String[] doc_words, Map<String, Color> highlight_colors){
		Map<String, String> hex_colors = new HashMap<String, String>();
		Set<String> keySet = highlight_colors.keySet();
		Object[] word_keys = keySet.toArray();
		String[] string_word_keys = Arrays.copyOf(word_keys, word_keys.length, String[].class);
		for(int i = 0; i < doc_words.length; i++){
			String[] temp_words = doc_words[i].split(" ");
			for(String word : temp_words){
				String temp_word = removePunctuation(word);
				temp_word = makeLowerCase(temp_word);
				if(Arrays.asList(string_word_keys).contains(temp_word)){
					Color color = highlight_colors.get(temp_word);
					String hex = String.format("#%02X%02X%02X", color.getRed(), color.getGreen(), color.getBlue()); 
					StringBuilder sb = new StringBuilder(hex);
					sb.deleteCharAt(0);
					String new_hex = sb.toString();
					hex_colors.put(word, new_hex);
				}
			}
		}
		return hex_colors;
	}
	
	public static String makeLowerCase(String word){
		String lowercase_word = "";
		for(int i = 0; i < word.length(); i++){
			if(Character.isUpperCase(word.charAt(i))){
				lowercase_word += Character.toLowerCase(word.charAt(i));
			}else{
				lowercase_word += word.charAt(i);
			}
		}
		return lowercase_word;
	}
	
	public static void appendExternalHyperlink(String url, String text, XWPFParagraph paragraph) throws InstantiationException, IllegalAccessException{

        
		
		//Add the link as External relationship
        String id = paragraph.getDocument().getPackagePart().addExternalRelationship(url, XWPFRelation.HYPERLINK.getRelation()).getId();

        //Append the link and bind it to the relationship
        org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHyperlink cLink = paragraph.getCTP().addNewHyperlink();
        cLink.setId(id);

        //Create the linked text
        CTText ctText=CTText.Factory.newInstance();
        ctText.setStringValue(text);
        CTR ctr=CTR.Factory.newInstance();
        
        //add blue shading to word and underline
        CTRPr rpr = ctr.addNewRPr(); 
        CTColor colour = CTColor.Factory.newInstance(); 
        colour.setVal("0000FF"); 
        rpr.setColor(colour); 
        CTRPr rpr1 = ctr.addNewRPr(); 
        rpr1.addNewU().setVal(STUnderline.SINGLE);
        
        ctr.setTArray(new CTText[]{ctText});

        //Insert the linked text into the link
        cLink.setRArray(new CTR[]{ctr});
        
    }
	
}
    	