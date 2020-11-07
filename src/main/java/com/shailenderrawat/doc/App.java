package com.shailenderrawat.doc;

import java.util.HashMap;
import java.util.List;

import javax.xml.bind.JAXBException;

import org.docx4j.Docx4J;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.opendope.answers.Answer;
import org.opendope.answers.Answers;
import org.opendope.answers.Repeat;
import org.opendope.answers.Repeat.Row;

/**
 * Hello world!
 */
public final class App {

    private static Answers answers;

    private App() {
    }

    /**
     * Says hello to the world.
     * 
     * @param args The arguments of the program.
     */
    public static void main(String[] args) {
        System.out.println("Hello World!");
        String inputfilepath = "D:\\shailender\\MyProjects\\EmployeeDetailsTemplate.docx";
        try {
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new java.io.File(inputfilepath));
            MainDocumentPart mdp = wordMLPackage.getMainDocumentPart();
    
            answers = new Answers();

            addAnswer("name_Yf", "John Smith");
            addAnswer("dob_Js", "01/10/2010"); // or false
            addAnswer("Nationality_da", "India"); // or false
            addAnswer("education_G4", "B. Tech"); // or false
            

            Repeat repeatRow = new Repeat();
            repeatRow.setQref("experience_y9");
            

            Row rowOne = new Row();

            Answer organization = new Answer();
            organization.setId("organization_OI");
            organization.setValue("DBS");
        
            Answer designation = new Answer();
            designation.setId("designation_S3");
            designation.setValue("VP");

            Answer experienceInCompany = new Answer();
            experienceInCompany.setId("experienceInCompany_kQ");
            experienceInCompany.setValue("2");

            rowOne.getAnswerOrRepeat().add(organization);
            rowOne.getAnswerOrRepeat().add(designation);
            rowOne.getAnswerOrRepeat().add(experienceInCompany);

            Row rowTwo = new Row();

            Answer organizationTwo = new Answer();
            organizationTwo.setId("organization_OI");
            organizationTwo.setValue("ABC Company");

        
            Answer designationTwo = new Answer();
            designationTwo.setId("designation_S3");
            designationTwo.setValue("Engineering Manager");

            Answer experienceInCompanyTwo = new Answer();
            experienceInCompanyTwo.setId("experienceInCompany_kQ");
            experienceInCompanyTwo.setValue("4");

            rowTwo.getAnswerOrRepeat().add(organizationTwo);
            rowTwo.getAnswerOrRepeat().add(designationTwo);
            rowTwo.getAnswerOrRepeat().add(experienceInCompanyTwo);

            
            repeatRow.getRow().add(rowOne);
            repeatRow.getRow().add(rowTwo);
            answers.getAnswerOrRepeat().add(repeatRow);

            String filename = System.getProperty("user.dir") + "\\OUT_hello_new.docx";

            Docx4J.bind(wordMLPackage, answers, Docx4J.FLAG_BIND_INSERT_XML | Docx4J.FLAG_BIND_BIND_XML);
            Docx4J.save(wordMLPackage, new java.io.File(filename), Docx4J.FLAG_NONE);

            System.out.println("Saved " + filename);

        } catch (Docx4JException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }

    private static void addAnswer(String key, String value) {
        Answer a = new Answer();
        a.setId(key);
        a.setValue(value);
        answers.getAnswerOrRepeat().add(a);
    }

}
