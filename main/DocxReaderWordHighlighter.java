package main;

import java.awt.Color;

import java.util.HashMap;
import java.util.Map;
import java.util.Random;
import java.util.Set;

public class DocxReaderWordHighlighter {

	private String[] doc_words_copy;
	private int user_word_highlight_num;
	private Map<String, Integer> words;
	private Map<String, Color> highlight_colors;
	
	public DocxReaderWordHighlighter(String[] doc_words_copy, int user_word_highlight_num){
		this.doc_words_copy = doc_words_copy;
		this.user_word_highlight_num = user_word_highlight_num;
		words = new HashMap<String, Integer>();
		highlight_colors = new HashMap<String, Color>();
	}
	
	public void putWordsInMap(){
		String paragraph;
		int string_location = 0;
		int space_location = 0;
		int previous_space = 0;
		for(int i = 0; i < doc_words_copy.length; i++){
			string_location = 0;
			paragraph = doc_words_copy[i];
			space_location = 0;
			previous_space = 0;
			while(string_location != -1){	
				if(string_location == 0){
					space_location = paragraph.indexOf(" ");
					//this condition ensures that if there are no spaces current loop is broken
					if(space_location == -1){
						String temp_word = paragraph;
						int z = temp_word.length();
						String other_temp_word = "";
					    for(int y = 0; y < z; y++){
					    	if(Character.isUpperCase(temp_word.charAt(y))){
					    		char w = temp_word.charAt(y);
					            other_temp_word += Character.toLowerCase(w);   
					         }else{
					        	 other_temp_word += temp_word.charAt(y);
					         }
					    }
					    temp_word = other_temp_word;
						if(!words.containsKey(temp_word)){
							words.put(temp_word, 1);
						}
						else{
							int temp_word_count = words.get(temp_word);
							temp_word_count++;
							words.put(temp_word, temp_word_count);
						}
						
						break;
					}
					else{	
						String temp_word = paragraph.substring(0, space_location);
						int z = temp_word.length();
						String other_temp_word = "";
					    for(int y = 0; y < z; y++){
					    	if(Character.isUpperCase(temp_word.charAt(y))){
					    		char w = temp_word.charAt(y);
					            other_temp_word += Character.toLowerCase(w);   
					         }else{
					        	 other_temp_word += temp_word.charAt(y);
					         }
					    }
					    temp_word = other_temp_word;
						if(!words.containsKey(temp_word)){
							words.put(temp_word, 1);
						}
						else{
							int temp_word_count = words.get(temp_word);
							temp_word_count++;
							words.put(temp_word, temp_word_count);
						}
						
						int temp_previous_space = space_location;
						int temp_space_location = paragraph.indexOf(" ", temp_previous_space + 1);
						if(temp_space_location == -1){
							string_location = 2;
						}
						else{
							string_location = 1;
						}
						
					}	
				}
				else if(string_location == 1){
					previous_space = space_location;
					space_location = paragraph.indexOf(" ", previous_space + 1);
					String temp_word = paragraph.substring(previous_space + 1, space_location);
					int z = temp_word.length();
					String other_temp_word = "";
				    for(int y = 0; y < z; y++){
				    	if(Character.isUpperCase(temp_word.charAt(y))){
				    		char w = temp_word.charAt(y);
				            other_temp_word += Character.toLowerCase(w);   
				         }else{
				        	 other_temp_word += temp_word.charAt(y);
				         }
				    }
				    temp_word = other_temp_word;
					if(!words.containsKey(temp_word)){
						words.put(temp_word, 1);
					}
					else{
						int temp_word_count = words.get(temp_word);
						temp_word_count++;
						words.put(temp_word, temp_word_count);
					}
					int temp_previous_space = space_location;
					int temp_space_location = paragraph.indexOf(" ", temp_previous_space + 1);
					if(temp_space_location == -1){
						string_location = 2;
					}
					
				}
				else{
					String temp_word = paragraph.substring(space_location + 1);
					int z = temp_word.length();
					String other_temp_word = "";
				    for(int y = 0; y < z; y++){
				    	if(Character.isUpperCase(temp_word.charAt(y))){
				    		char w = temp_word.charAt(y);
				            other_temp_word += Character.toLowerCase(w);   
				         }else{
				        	 other_temp_word += temp_word.charAt(y);
				         }
				    }
				    temp_word = other_temp_word;
					if(!words.containsKey(temp_word)){
						words.put(temp_word, 1);
					}
					else{
						int temp_word_count = words.get(temp_word);
						temp_word_count++;
						words.put(temp_word, temp_word_count);
					}
					string_location = -1;
					
				}
			}	
		}
		highlightWords();
	}
	
	public void highlightWords(){
		Set<String> keySet = words.keySet();
		Object[] word_keys = keySet.toArray();
		for(int i = 0; i < words.size(); i++){
			//make sure string isn't just whitespace here
			String temp = word_keys[i].toString();
			if(!isCommonWord(temp)){
				if(!temp.trim().isEmpty()){
					if(words.get(word_keys[i]) >= user_word_highlight_num){
						Color randomColor = generateRandomColor();
						highlight_colors.put(word_keys[i].toString(), randomColor);		
					}
				}	
			}	
		}
	}
	
	public Map<String, Color> getHighlightColors(){
		return highlight_colors;
	}
	
	public Color generateRandomColor(){
		Random rand = new Random();
		Color randomColor;
		int r;
		int g;
		int b;
		int rand_num;
		do{	
			r = rand.nextInt(256);
			g = rand.nextInt(256);
			b = rand.nextInt(256);
			rand_num = rand.nextInt(3);
			if(rand_num == 0){
				r = rand.nextInt(31);
			}
			else if(rand_num == 1){
				b = rand.nextInt(31);
			}
			else{
				g = rand.nextInt(31);
			}
			randomColor = new Color(r, g, b);	
		}while(highlight_colors.containsKey(randomColor) || areColorsBland(r, g, b) == false);
		
		return randomColor;
	}
	
	public boolean areColorsBland(int r, int g, int b){
		if((r >= 200 && g <= 35 && b <= 35) || (b >= 200 && g <= 35 && r <= 35) ||
				(g >= 200 && r <= 35 && b <= 35)){
			return true;
		}
		else if((r >= 200 && b >= 100 && b <= 200 && g <= 75) || (b >= 200 && r >= 100
				&& r <= 200 && g <= 75) || (g >= 200 && b >= 100 && b <= 200 && r <= 75) ||
				(g >= 200 && r >= 100 && r <= 200 && b <= 75) || (b >= 200 && g >= 100 && g <= 200 
				&& r <= 75) || (r >= 200 && g >= 100 && g <= 200 && b <= 75)){
			return true;
		}
		else{
			return false;
		}	
	}
	
	public boolean isCommonWord(String word){
		switch(word){
		
			case "this": return true;
			case "a": return true;
			case "the": return true;
			case "was": return true;
			case "and": return true;
			case "how": return true;
			case "to": return true;
			case "for": return true;
			case "what": return true;
			case "who": return true;
			case "when": return true;
			case "where": return true;
			case "in": return true;
			case "on": return true;
			case "of": return true;
			case "is": return true;
			case "that": return true;
			case "as": return true;
			
		}
		
		return false;
	}
	
	
}
