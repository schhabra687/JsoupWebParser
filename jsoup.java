

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

public class jsoup {
	
	static int count = 0;
	static int row = 0;
	static int column = 0;
	static int flag = 0;
	
	final static String memberId1 = "Membership ID";
	final static String companyName1 = "Company Name";
	final static String contactPerson1 = "Contact Person";
	//final static String chapter1 = "Chapter:";
	final static String referencePerson1 = "Reference Person";
	final static String address1 = "Address";
	//final static String district1 = "District:";
	final static String pinCode1 = "Pin Code";
	final static String mobile1 = "Mobile Numbers";
	final static String state1 = "State";
	final static String stdCode1 = "STD Code";
	final static String phoneNum1 = "Phone Numbers";
	//final static String country1 = "Country";
	final static String phoneResi1 = "Residence Phone";
	final static String email1 = "E Mail Address";
	//final static String phoneOffice1 = "Phone Office:";
	final static String website1 = "Website";
	final static String fax1 = "Fax Numbers";
	//final static String category1 = "Category:";
	//final static String subCategory1 = "Sub Category :";
	//final static String dealsIn1 = "Deals In:";
	final static String natureOfBusiness1 = "Nature Of Business";
	final static String productDetails1 = "Product Details";
	
	
	static String memberId = "";
	static String companyName = "";
	static String contactPerson = "";
	static String referencePerson = "";
	//static String chapter = "";
	static String address = "";
	static String pinCode = "";
	//static String district = "";
	static String mobile = "";
	static String state = "";
	static String stdCode = "";
	static String phoneNum = "";
	static String phoneResi = "";
	//static String country = "";
	//static String phoneFact = "";
	static String email = "";
	//static String phoneOffice = "";
	static String website = "";
	static String fax = "";
	//static String category = "";
	//static String subCategory = "";
	//static String dealsIn = "";
	static String natureOfBusiness = "";
	static String productDetails = "";
		
	
	static Map indexMap = new HashMap();
    static List<Integer> indexList=new ArrayList();
    static Map detailMap = new HashMap();
	static List<Integer> detailList=new ArrayList();
	
	public static void main(String [] args) throws Exception{
		
		XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Java Books");
        //extract("smeoutput1",workbook,sheet);
        for(char i='b';i<='b';++i){
        	flag = 0;
		extract(""+i,workbook,sheet);
		System.out.println("Entry for file count"+i+" written successfully..");
        }
        
        /*
        for(int i=500;i<=578;i++){
        	flag = 0;
		extract("output"+i,workbook,sheet);
        }*/
		try (FileOutputStream outputStream = new FileOutputStream("C:\\Users\\Shubham-PC\\Downloads\\test.xlsx")) {
            workbook.write(outputStream);
        }
		
	}

	public static void extract (String htmlName,XSSFWorkbook workbook,XSSFSheet sheet) throws Exception {
	    
		File input = new File("C:\\Users\\Shubham-PC\\Downloads\\"+htmlName+".html");
		Document doc=null;
		try {
			doc = Jsoup.parse(input, "UTF-8", "http://test.com/");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		int k=0;
		Elements link = doc.select("table[class=MsoNormalTable]");
		
		//System.out.println("data:"+link.text());
		for(Element div :link){
			System.out.println(htmlName+k++);
			String name = div.text();
			//String name = link.text();
			System.out.println(name);
			parseInfo(name,workbook,sheet);
		}
		
		
		 //System.out.println(" start :  ");
		
		 //Elements link = doc.getElementsByTag("table");
		 //String name = link.text();
		// Element table = doc.select("table[border=0]").first();
		 //System.out.println(table.text());
		 //String name = table.text();
		 //System.out.println(name);		
		 //parseInfo(name+17,workbook,sheet);
		// parseInfo(name,workbook,sheet);		 
		 //System.out.println("end");
	}
	
		
	public static void write (XSSFWorkbook workbook,XSSFSheet sheet) {
		
		/* int index = email.indexOf(",");
		 if(index!=-1)
			 email = email.substring(0, index);
		*/ 
                  
		 Row row1 = sheet.createRow(row++);		
		 
		 memberId=memberId.trim();
		 companyName=companyName.trim();
		 contactPerson=contactPerson.trim();
		 referencePerson=referencePerson.trim();
		 //chapter=chapter.trim();
		 address=address.trim();
		 pinCode=pinCode.trim();
		 //district=district.trim();
		 mobile=mobile.trim();
		 state=state.trim();
		 stdCode=stdCode.trim();
		 phoneNum=phoneNum.trim();
		 phoneResi=phoneResi.trim();
		 //country=country.trim();
		 //phoneFact=phoneFact.trim();
		 email=email.trim();
		 //phoneOffice=phoneOffice.trim();
		 website=website.trim();
		 fax=fax.trim();
		 //category=category.trim();
		 //subCategory=subCategory.trim();
		 //dealsIn=dealsIn.trim();
		 natureOfBusiness=natureOfBusiness.trim();
		 productDetails=productDetails.trim();
			
			
		 row1.createCell(0).setCellValue(memberId);
		 row1.createCell(1).setCellValue(companyName);
		 row1.createCell(2).setCellValue(contactPerson);
		 row1.createCell(3).setCellValue(referencePerson);
		 //row1.createCell(3).setCellValue(chapter);
		 row1.createCell(4).setCellValue(address);
		 row1.createCell(5).setCellValue(pinCode);
		 //row1.createCell(5).setCellValue(district);
		 row1.createCell(7).setCellValue(mobile);
		 row1.createCell(6).setCellValue(state);
		 row1.createCell(8).setCellValue(stdCode);
		 row1.createCell(9).setCellValue(phoneNum);
		 row1.createCell(11).setCellValue(phoneResi);
		 //row1.createCell(9).setCellValue(country);
		 //row1.createCell(10).setCellValue(phoneFact);
		 row1.createCell(12).setCellValue(email);
		 //row1.createCell(12).setCellValue(phoneOffice);
		 row1.createCell(13).setCellValue(website);
		 row1.createCell(10).setCellValue(fax);
		 //row1.createCell(15).setCellValue(category);
		 //row1.createCell(16).setCellValue(subCategory);
		 //row1.createCell(17).setCellValue(dealsIn);
		 row1.createCell(14).setCellValue(natureOfBusiness);
		 row1.createCell(15).setCellValue(productDetails);
	}
	
	public static String getSubstringValue(int startIndex,String fieldName,List fieldList,String subFieldDetails){
		String value="";
		int startValue=0;
		if(fieldName.contains("\\n")){
			startValue=startIndex+fieldName.length();
		}
		else{
			startValue=startIndex+fieldName.length();
		}
		
		int endValue=getIndexOfValue(fieldName, fieldList, subFieldDetails);
		
		int length=endValue-startValue;
		if((startIndex!=-1)&&(length>4)){
			value=subFieldDetails.substring(startValue,endValue);
		}
		return value;
	}
	
	public static int getStartIndex(String subFieldDetails,String fieldName,int flag){
		int startIndex=subFieldDetails.indexOf(fieldName);
		//System.out.println(startIndex);
		if(startIndex!=-1){
			if(flag==1){
				indexList.add(startIndex);
				indexMap.put(startIndex,fieldName);
			}
			else{
				detailList.add(startIndex);
				detailMap.put(startIndex,fieldName);
			}
		}
		flag=0;
       return startIndex;  
	}
	
	public static int getIndexOfValue(String str,List fieldList, String mainString){
		int length=0;
		for (int i=0; i<fieldList.size(); i++) {
			length=fieldList.get(i).toString().length();
			if(fieldList.get(i).equals(str) && i!=fieldList.size()-1) {
				length = mainString.indexOf((String)fieldList.get(i+1));
				return length;
			}
		}
		
		return mainString.length()-59;
	}
	
	private static void parseInfo(String line,XSSFWorkbook workbook,XSSFSheet sheet) throws Exception {
		
	 		String subFieldDetails = line;
	 		
	 		 try {
        		
                List fieldList=new ArrayList();         
                int flag=1;

                
                int memberIdStartIndex = getStartIndex(subFieldDetails,memberId1,flag);
                int companyNameStartIndex =  getStartIndex(subFieldDetails,companyName1,flag);
                int contactPersonStartIndex = getStartIndex(subFieldDetails,contactPerson1,flag);
                int referencePersonStartIndex = getStartIndex(subFieldDetails,referencePerson1,flag);
                //int chapterStartIndex = getStartIndex(subFieldDetails,chapter1,flag);
                int addressStartIndex = getStartIndex(subFieldDetails,address1,flag);
                //int districtStartIndex = getStartIndex(subFieldDetails,district1,flag);
                int mobileStartIndex = getStartIndex(subFieldDetails,mobile1,flag);
                int stateStartIndex = getStartIndex(subFieldDetails,state1,flag);
                int phoneResiStartIndex = getStartIndex(subFieldDetails,phoneResi1,flag);
                //int countryStartIndex = getStartIndex(subFieldDetails,country1,flag);
                //int phoneFactStartIndex = getStartIndex(subFieldDetails,phoneFact1,flag);
                int emailStartIndex = getStartIndex(subFieldDetails,email1,flag);
                //int phoneOfficeStartIndex = getStartIndex(subFieldDetails,phoneOffice1,flag);
                int websiteStartIndex = getStartIndex(subFieldDetails,website1,flag);
                int faxStartIndex = getStartIndex(subFieldDetails,fax1,flag);
                //int categoryStartIndex = getStartIndex(subFieldDetails,category1,flag);
                //int subCategoryStartIndex = getStartIndex(subFieldDetails,subCategory1,flag);
                //int dealsInStartIndex = getStartIndex(subFieldDetails,dealsIn1,flag);
                int pinCodeStartIndex = getStartIndex(subFieldDetails,pinCode1,flag);
                int stdCodeStartIndex = getStartIndex(subFieldDetails,stdCode1,flag);
                int phoneNumStartIndex = getStartIndex(subFieldDetails,phoneNum1,flag);
                int natureOfBusinessStartIndex = getStartIndex(subFieldDetails,natureOfBusiness1,flag);
                int productDetailsStartIndex = getStartIndex(subFieldDetails,productDetails1,flag);
                
                Collections.sort(indexList);
                for(int i=0; i<indexList.size(); i++) {
                	fieldList.add(indexMap.get(indexList.get(i)));
                }
               
                memberId=getSubstringValue(memberIdStartIndex,memberId1, fieldList, subFieldDetails);
                companyName=getSubstringValue(companyNameStartIndex,companyName1, fieldList, subFieldDetails);
                contactPerson=getSubstringValue(contactPersonStartIndex,contactPerson1, fieldList, subFieldDetails);
                referencePerson=getSubstringValue(referencePersonStartIndex,referencePerson1, fieldList, subFieldDetails);
                //chapter=getSubstringValue(chapterStartIndex,chapter1, fieldList, subFieldDetails);
                address=getSubstringValue(addressStartIndex,address1, fieldList, subFieldDetails);
                //district=getSubstringValue(districtStartIndex,district1, fieldList, subFieldDetails);
                mobile=getSubstringValue(mobileStartIndex,mobile1, fieldList, subFieldDetails);
                state=getSubstringValue(stateStartIndex,state1, fieldList, subFieldDetails);
                phoneResi=getSubstringValue(phoneResiStartIndex,phoneResi1, fieldList, subFieldDetails);
                //country=getSubstringValue(countryStartIndex,country1, fieldList, subFieldDetails);
                //phoneFact=getSubstringValue(phoneFactStartIndex,phoneFact1, fieldList, subFieldDetails);
                email=getSubstringValue(emailStartIndex,email1, fieldList, subFieldDetails);
                //phoneOffice=getSubstringValue(phoneOfficeStartIndex,phoneOffice1, fieldList, subFieldDetails);
                website=getSubstringValue(websiteStartIndex,website1, fieldList, subFieldDetails);
                fax=getSubstringValue(faxStartIndex,fax1, fieldList, subFieldDetails);
                //category=getSubstringValue(categoryStartIndex,category1, fieldList, subFieldDetails);
                //subCategory=getSubstringValue(subCategoryStartIndex,subCategory1, fieldList, subFieldDetails);
                //dealsIn=getSubstringValue(dealsInStartIndex,dealsIn1, fieldList, subFieldDetails);
                pinCode=getSubstringValue(pinCodeStartIndex,pinCode1, fieldList, subFieldDetails);
                stdCode=getSubstringValue(stdCodeStartIndex,stdCode1, fieldList, subFieldDetails);
                phoneNum=getSubstringValue(phoneNumStartIndex,phoneNum1, fieldList, subFieldDetails);
                natureOfBusiness=getSubstringValue(natureOfBusinessStartIndex,natureOfBusiness1, fieldList, subFieldDetails);
                productDetails=getSubstringValue(productDetailsStartIndex,productDetails1, fieldList, subFieldDetails);
                
                /* System.out.println(memberId);
                System.out.println(companyName);
                System.out.println(contactPerson);
                System.out.println(chapter);
    			System.out.println(address);
    			System.out.println(district);
    			System.out.println(mobile);
    			System.out.println(state);
    			System.out.println(phoneResi);
    			System.out.println(country);
    			System.out.println(phoneFact);
    			//System.out.println(phone);
    			System.out.println(fax);
    			System.out.println(email);
    			System.out.println(phoneOffice);
    			System.out.println(website);
    			System.out.println(category);
    			System.out.println(subCategory);
    			System.out.println(dealsIn);*/
    			//System.out.println(type);
  		
    			write (workbook,sheet);
    			indexList=new ArrayList();
            			indexMap=new HashMap();           
 	 		 }
	        catch(Exception e){
	        	
	        }
	}
}
