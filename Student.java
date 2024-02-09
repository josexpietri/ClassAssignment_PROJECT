package CodingProject;
import java.io.*;
import java.io.FileInputStream;
import java.util.Scanner;


public class Student {

String firstname;
String lastname;
String p1;
String p2;
String p3;
String p4;
String p5;
int schedule;


    public static void main(String[] args) {
        

        try{

            Scanner s1 = new Scanner(System.in);
            System.out.println("Enter File Name: (Preference Excel Sheet Name)");
            //ENTER THE NAME OF THE EXCEL SHEET WITH THE STUDENTS PREFERENCES
            String filepath1 = s1.next(); 

            File file = new File("C:\\SeniorScheduling\\"+filepath1+".xlsx");
            FileInputStream fO = new FileInputStream(file); 
            XSSFWorkbook wb = new XSSFWorkbook(fO);
            XSSFSheet sheet = wb.getSheetAt(0);
            int numberofstudents = sheet.getPhysicalNumberOfRows();
            Student student[] = new Student[numberofstudents];

            //READ THE PREFERENCES OF THE STUDENTS
            readExcel(student, numberofstudents, sheet);
             //Students with p1,p2,p3 being LANGUAGE - PHYSICS - CALCULUS   [in any order]
         Prio prio1[] = new Prio[numberofstudents];

         //lengths of how many students per language that are prio1
         int lengths[] = assignPrio1(student, numberofstudents, prio1);
 
         String prio1S[] = new String[lengths[1]]; //Prio1 Spanish
         String prio1F[] = new String[lengths[2]]; //Prio1 French
         String prio1G[] = new String[lengths[3]]; //Prio1 German
         String prio1L[] = new String[lengths[4]]; //Prio1 Greek
 
         //assigning the studnets with the prio1 preferences to their category according to their language
         assignS(prio1, prio1S, student, prio1F, prio1G, prio1L, lengths);
     
         //Prio2 Students - students with other combinations of preferences
         Prio prio2[] = new Prio[numberofstudents];
         int lengthp2 =  assignPrio2(prio2, student, prio1, numberofstudents, lengths); //gives prio2 length + assigns stduents to prio2
 
         //ALL POSSIBLE SUBJECTS
         String physics = "Physics II";
         String calc = "Calculus II";
         String rhet = "American Rhetorical Tradition";
         String code = "Logic and Coding";
         String spa = "Language - Spanish IV";
         int s = 0;
 
         String fre = "Language - French IV";
         int f = 1;
 
         String ger = "Language - German IV";
         int g = 2;
 
         String greek = "Language - Greek II";
         int l = 3;
 
         //Array of Prefrence --- ie LRCo = Language - Rhetoric - Coding
         //int LLLS >> Language = Spanish [0]
         //int LLLF >> Language = French [1]
         //int LLLG >> Language = German [2]
         //int LLLL >> Language = Greek [3]
 
 
         //LRCo = Language - Rhetoric - Coding
         String prio2LRCo[][] = new String[4][lengthp2];
         int prio2LRCoLLLS = checkP2Students(student, prio2LRCo, numberofstudents, spa, rhet, code,s);
         int prio2LRCoLLLF = checkP2Students(student, prio2LRCo, numberofstudents, fre, rhet, code,f);
         int prio2LRCoLLLG = checkP2Students(student, prio2LRCo, numberofstudents, ger, rhet, code,g);
         int prio2LRCoLLLL = checkP2Students(student, prio2LRCo, numberofstudents, greek, rhet, code,l);
 
         //LRCa = Language - Rhetoric - Calculus
         String prio2LRCa[][] = new String[4][lengthp2];
         int prio2LRCaLLLS = checkP2Students(student, prio2LRCa, numberofstudents, spa, rhet, calc,s);
         int prio2LRCaLLLF = checkP2Students(student, prio2LRCa, numberofstudents, fre, rhet, calc,f);
         int prio2LRCaLLLG = checkP2Students(student, prio2LRCa, numberofstudents, ger, rhet, calc,g);
         int prio2LRCaLLLL = checkP2Students(student, prio2LRCa, numberofstudents, greek, rhet, calc,l);
 
         //LRP = Language - Rhetoric - Physics
         String prio2LRP[][] = new String[4][lengthp2];
         int prio2LRPLLLS = checkP2Students(student, prio2LRP, numberofstudents, spa, rhet, physics,s);
         int prio2LRPLLLF = checkP2Students(student, prio2LRP, numberofstudents, fre, rhet, physics,f);
         int prio2LRPLLLG = checkP2Students(student, prio2LRP, numberofstudents, ger, rhet, physics,g);
         int prio2LRPLLLL = checkP2Students(student, prio2LRP, numberofstudents, greek, rhet, physics,l);
 
         //LCaCo = Language - Calculus - Coding
         String prio2LCaCo[][] = new String[4][lengthp2];
         int prio2LCaCoLLLS = checkP2Students(student, prio2LCaCo, numberofstudents, spa, calc, code,s);                              
         int prio2LCaCoLLLF = checkP2Students(student, prio2LCaCo, numberofstudents, fre, calc, code,f);
         int prio2LCaCoLLLG = checkP2Students(student, prio2LCaCo, numberofstudents, ger, calc, code,g);
         int prio2LCaCoLLLL = checkP2Students(student, prio2LCaCo, numberofstudents, greek, calc, code,l);
 
         //LCaCo = Language - Coding - Physics
         String prio2LCoP[][] = new String[4][lengthp2];
         int prio2LCoPLLLS = checkP2Students(student, prio2LCoP, numberofstudents, spa, code, physics,s);
         int prio2LCoPLLLF = checkP2Students(student, prio2LCoP, numberofstudents, fre, code, physics,f);
         int prio2LCoPLLLG = checkP2Students(student, prio2LCoP, numberofstudents, ger, code, physics,g);
         int prio2LCoPLLLL = checkP2Students(student, prio2LCoP, numberofstudents, greek, code, physics,l);
 
         //LCaCo = Physics - Calculus - Rhetoric
         String prio2PCaR[] = new String[lengthp2];
         int prio2PCaRLLL = checkP2P2Students(student, prio2PCaR, numberofstudents, physics, calc, rhet);
 
         //LCaCo = Physics - Calculus - Coding
         String prio2PCaCo[] = new String[lengthp2];
         int prio2PCaCoLLL = checkP2P2Students(student, prio2PCaCo, numberofstudents, physics, calc, code);
         
         //LCaCo = Physics - Rhetoric - Coding
         String prio2PRCo[] = new String[lengthp2];
         int prio2PRCoLLL = checkP2P2Students(student, prio2PRCo, numberofstudents, physics, code, rhet);
     
         //LCaCo = Coding - Rhetoric - Calculus   
         String prio2CoRCa[] = new String[lengthp2];
         int prio2CoRCaLLL = checkP2P2Students(student, prio2CoRCa, numberofstudents, code, calc, rhet);
        
 
         //NAME OF EXCEL SHEET WHERE THE SCHEDULES WILL BE PRINTED
         System.out.println("Enter File Path: (where the schedules will print)");
         String filepath = s1.next();
 
         
         //w1 = row of excel sheet
         int w1 = 0;
 
         w1 =  print1stLine(w1, filepath);
         w1 = printSchedulesP1( prio1S, prio1F, prio1G, prio1L, numberofstudents, filepath);
         int y = 2;
         w1  = printSP2p1( prio2LRCo, prio2LRCoLLLS, prio2LRCoLLLF, prio2LRCoLLLG, prio2LRCoLLLL, numberofstudents, y, filepath, w1);
         y = 3;
         w1 = printSP2p1( prio2LRCa, prio2LRCaLLLS, prio2LRCaLLLF, prio2LRCaLLLG, prio2LRCaLLLL, numberofstudents, y, filepath, w1);
         y = 4;
         w1 = printSP2p1( prio2LRP, prio2LRPLLLS, prio2LRPLLLF, prio2LRPLLLG, prio2LRPLLLL, numberofstudents, y, filepath, w1);
         y = 5;
         w1 = printSP2p1( prio2LCaCo, prio2LCaCoLLLS, prio2LCaCoLLLF, prio2LCaCoLLLG, prio2LCaCoLLLL, numberofstudents, y, filepath, w1);
         y = 6;
         w1 = printSP2p1( prio2LCoP, prio2LCoPLLLS, prio2LCoPLLLF, prio2LCoPLLLG, prio2LCoPLLLL, numberofstudents, y, filepath, w1);
         y = 7;
         w1 = printP2p2( prio2PCaR, prio2PCaRLLL, y, filepath, w1);
         y = 8;
         w1 = printP2p2( prio2PCaCo, prio2PCaCoLLL, y, filepath, w1);
         y = 9;
         w1 = printP2p2( prio2PRCo, prio2PRCoLLL, y, filepath, w1);
         y = 10;
         w1 = printP2p2( prio2CoRCa, prio2CoRCaLLL, y, filepath, w1);
 
         

            
            System.out.println("Done");
           wb.close();
           s1.close();

        }catch(Exception e){
            e.printStackTrace();
        }

    }

        
    public static void readExcel(Student student[], int numberofstudents, XSSFSheet sheet){
        
            for(int i=0;i<numberofstudents;i++){

                student[i] = new Student();
    
                //READING the row of prefernecs for each student
                XSSFRow row1 = sheet.getRow(i);

                //LAST NAME
                XSSFCell cellA1 = row1.getCell((short) 0);
                String a1Val = cellA1.getStringCellValue();
                //FIRST NAME
                XSSFCell cellB1 = row1.getCell((short) 1);
                String b1Val = cellB1.getStringCellValue();
                //PREF 1
                XSSFCell cellC1 = row1.getCell((short) 2);
                String c1Val = cellC1.getStringCellValue();
                //PREF 2
                XSSFCell cellD1 = row1.getCell((short) 3);
                String d1Val = cellD1.getStringCellValue();
                //PREF 3
                XSSFCell cellE1 = row1.getCell((short) 4);
                String e1Val = cellE1.getStringCellValue();
                //PREF 4
                XSSFCell cellF1 = row1.getCell((short) 5);
                String f1Val = cellF1.getStringCellValue();
                //PREF 5
                XSSFCell cellG1 = row1.getCell((short) 6);
                String g1Val = cellG1.getStringCellValue();
               
                
                //ASSIGNING PREFERENCES TO EACH STUDENT to place in category latet
                student[i].lastname = b1Val;
                student[i].firstname = a1Val;
                student[i].p1 = c1Val;
                student[i].p2 = d1Val;
                student[i].p3 = e1Val;
                student[i].p4 = f1Val;
                student[i].p5 = g1Val;
               
    
            }
        }

   

   

   

    public static int[] assignPrio1(Student student[], int numberofstudents, Prio prio1[]){

        int lengths[] = new int[5];
            lengths[0] = 0;
            lengths[1] = 0;
            lengths[2] = 0;
            lengths[3] = 0;
            lengths[4] = 0;


            for(int i=0;i<numberofstudents;i++){

            
                String name = student[i].firstname +" "+ student[i].lastname;
    
                //Spanish
               
                
                  if ((student[i].p1.equals("Physics II") && student[i].p2.equals("Language - Spanish IV") && student[i].p3.equals("Calculus II")) ||
                (student[i].p1.equals("Physics II") && student[i].p2.equals("Calculus II") && student[i].p3.equals("Language - Spanish IV")) ||
                (student[i].p1.equals("Calculus II") && student[i].p2.equals("Language - Spanish IV") && student[i].p3.equals("Physics II"))||
                (student[i].p1.equals("Calculus II") && student[i].p2.equals("Physics II") && student[i].p3.equals("Language - Spanish IV"))||
                (student[i].p1.equals("Language - Spanish IV") && student[i].p2.equals("Calculus II") && student[i].p3.equals("Physics II"))||
                (student[i].p1.equals("Language - Spanish IV") && student[i].p2.equals("Physics II") && student[i].p3.equals("Calculus II"))) {
                    
                    prio1[lengths[0]] = new Prio();
                    prio1[lengths[0]].l = "Language - Spanish IV";
                    prio1[lengths[0]].stu = i;
                    prio1[lengths[0]].name = name;
                    lengths[0]++;
                    lengths[1]++;
            
                 }else 
                 
                 
                 if ((student[i].p1.equals("Physics II") && student[i].p2.equals("Language - French IV") && student[i].p3.equals("Calculus II")) ||
                 (student[i].p1.equals("Physics II") && student[i].p2.equals("Calculus II") && student[i].p3.equals("Language - French IV")) ||
                 (student[i].p1.equals("Calculus II") && student[i].p2.equals("Language - French IV") && student[i].p3.equals("Physics II"))||
                 (student[i].p1.equals("Calculus II") && student[i].p2.equals("Physics II") && student[i].p3.equals("Language - French IV"))||
                 (student[i].p1.equals("Language - French IV") && student[i].p2.equals("Calculus II") && student[i].p3.equals("Physics II"))||
                 (student[i].p1.equals("Language - French IV") && student[i].p2.equals("Physics II") && student[i].p3.equals("Calculus II"))) {
     
                     prio1[lengths[0]] = new Prio();
                     prio1[lengths[0]].l = "Language - French IV";
                     prio1[lengths[0]].stu = i;
                     prio1[lengths[0]].name = name;
                     lengths[0]++;
                     lengths[2]++;
             
                 }else 
                 
                 if ((student[i].p1.equals("Physics II") && student[i].p2.equals("Language - German IV") && student[i].p3.equals("Calculus II")) ||
                (student[i].p1.equals("Physics II") && student[i].p2.equals("Calculus II") && student[i].p3.equals("Language - German IV")) ||
                (student[i].p1.equals("Calculus II") && student[i].p2.equals("Language - German IV") && student[i].p3.equals("Physics II"))||
                (student[i].p1.equals("Calculus II") && student[i].p2.equals("Physics II") && student[i].p3.equals("Language - German IV"))||
                (student[i].p1.equals("Language - German IV") && student[i].p2.equals("Calculus II") && student[i].p3.equals("Physics II"))||
                (student[i].p1.equals("Language - German IV") && student[i].p2.equals("Physics II") && student[i].p3.equals("Calculus II"))) {
    
                    prio1[lengths[0]] = new Prio();
                    prio1[lengths[0]].l = "Language - German IV";
                    prio1[lengths[0]].stu = i;
                    prio1[lengths[0]].name = name;
                    lengths[0]++;
                    lengths[3]++;
            
                 }else
                 
                 
    
                 
                 if ((student[i].p1.equals("Physics II") && student[i].p2.equals("Language - Greek II") && student[i].p3.equals("Calculus II")) ||
                (student[i].p1.equals("Physics II") && student[i].p2.equals("Calculus II") && student[i].p3.equals("Language - Greek II")) ||
                (student[i].p1.equals("Calculus II") && student[i].p2.equals("Language - Greek II") && student[i].p3.equals("Physics II"))||
                (student[i].p1.equals("Calculus II") && student[i].p2.equals("Physics II") && student[i].p3.equals("Language - Greek II"))||
                (student[i].p1.equals("Language - Greek II") && student[i].p2.equals("Calculus II") && student[i].p3.equals("Physics II"))||
                (student[i].p1.equals("Language - Greek II") && student[i].p2.equals("Physics II") && student[i].p3.equals("Calculus II"))) {
    
                    prio1[lengths[0]] = new Prio();
                    prio1[lengths[0]].l = "Language - Greek II";
                    prio1[lengths[0]].stu = i;
                    prio1[lengths[0]].name = name;
                    lengths[0]++;
                    lengths[4]++;
                
            
                 }
                
                
            }
        
          return lengths;         


        }
    
    public static void assignS(Prio prio1[], String prio1S[], Student student[], String prio1F[], String prio1G[], String prio1L[], int lengths[]){

            int x = 0;
            int l = 0;
            int g = 0;
            int f = 0;

            for(int i=0;i<lengths[0];i++){

                String name = student[prio1[i].stu].firstname +" "+ student[prio1[i].stu].lastname;

                if(prio1[i].l.equals("Language - Spanish IV")){
                    prio1S[x] = name;
                    x++;
                }else  if(prio1[i].l.equals("Language - French IV")){
                    prio1F[f] = name;
                    f++;
                }else  if(prio1[i].l.equals("Language - German IV")){
                    prio1G[g] = name;
                    g++;
                }else  if(prio1[i].l.equals("Language - Greek II")){
                    prio1L[l] = name;
                    l++;
                }

            }
        }
    
    public static int assignPrio2(Prio prio2[], Student student[], Prio prio1[], int numberofstudents, int lengths[]){

            int length = 0;

            for(int i=0;i<numberofstudents;i++){

                int x = 0;
                String name = student[i].firstname +" "+ student[i].lastname;

                for(int j=0;j<lengths[0];j++){
                
                    if(name.equals(prio1[j].name)){
                        x++;
                    }
                }



                if(x == 0){
                    
                    prio2[length] = new Prio();
                    prio2[length].name = name;
                    length++;
                }
            }

            return length;
        }
    
    
    public static int checkP2Students(Student student[], String pArray[][], int length, String S1, String S2, String S3, int x){

        int num = 0;


        for(int i=0;i<length;i++){

           String name = student[i].firstname +" "+ student[i].lastname;

               
            if((student[i].p1.equals(S1) && student[i].p2.equals(S2) && student[i].p3.equals(S3)) ||
                   (student[i].p1.equals(S1)&& student[i].p2.equals(S3)&& student[i].p3.equals(S2))||
                   (student[i].p1.equals(S3)&& student[i].p2.equals(S1)&& student[i].p3.equals(S2))||
                   (student[i].p1.equals(S3)&& student[i].p2.equals(S2)&& student[i].p3.equals(S1))||
                   (student[i].p1.equals(S2)&& student[i].p2.equals(S1)&& student[i].p3.equals(S3))||
                   (student[i].p1.equals(S2)&& student[i].p2.equals(S3)&& student[i].p3.equals(S1))){


                       pArray[x][num] = name;
                       num++;
                       


                   }
        }

           return num;
   }

       
    public static int checkP2P2Students(Student student[], String pArray[], int length, String S1, String S2, String S3){

        int num = 0;


        for(int i=0;i<length;i++){

           String name = student[i].firstname +" "+ student[i].lastname;

               
            if((student[i].p1.equals(S1) && student[i].p2.equals(S2) && student[i].p3.equals(S3)) ||
                   (student[i].p1.equals(S1)&& student[i].p2.equals(S3)&& student[i].p3.equals(S2))||
                   (student[i].p1.equals(S3)&& student[i].p2.equals(S1)&& student[i].p3.equals(S2))||
                   (student[i].p1.equals(S3)&& student[i].p2.equals(S2)&& student[i].p3.equals(S1))||
                   (student[i].p1.equals(S2)&& student[i].p2.equals(S1)&& student[i].p3.equals(S3))||
                   (student[i].p1.equals(S2)&& student[i].p2.equals(S3)&& student[i].p3.equals(S1))){


                       pArray[num] = name;
                       num++;
                       


                   }
        }

           return num;

       }
    

    public static int printSchedulesP1(String[] prio1S, String[] prio1F, String[] prio1G, String[] prio1L, int numberofstudents, String filepath){

      
            

            int w1 = 1;

            int x = 1;


            w1 = splitSection(w1, filepath);
            for(int j = 0;j<prio1S.length;j++){

                int y = 1;
               
               printonE2(x, w1, prio1S[j], y, filepath);
                
                
                w1++;

                }
            
                w1 = splitSection(w1, filepath);
            for(int j = 0;j<prio1F.length;j++){

                int y = 2;
                printonE2(x, w1, prio1F[j], y, filepath);
                
                
                w1++;
               

            }
            
            w1 = splitSection(w1, filepath);
            for(int j = 0;j<prio1G.length;j++){

                int y = 3;
                printonE2(x, w1, prio1G[j], y, filepath);
                
                
                w1++;
               

            }
             
            w1 = splitSection(w1, filepath);
            for(int j = 0;j<prio1L.length;j++){

                int y = 4;
                printonE2(x, w1, prio1L[j], y, filepath);
                
                
                w1++;

            }

            

            return w1;

       
    }

    
    public static int printSP2p1(String[][] array, int Slen, int Flen, int Glen, int Llen, int numberofstudents, int y, String filepath, int w1){

      
            
            
            int g = 0;
    
            
            w1 = splitSection(w1, filepath);
    
                for(int j = 0;j<Slen;j++){
    
                        g = 1;
                        printonE2(y, w1, array[0][j], g, filepath);
    
                        w1++;
                   
    
                    
                }
            
             w1 = splitSection(w1, filepath);
                for(int j = 0;j<Flen;j++){
    
                        g = 2;
                        printonE2(y, w1, array[1][j], g, filepath);
    
                        w1++;
                   
                }
                w1 = splitSection(w1, filepath);
                for(int j = 0;j<Glen;j++){
    
                    g = 3;
                    printonE2(y, w1, array[2][j], g, filepath);

                    w1++;
                   
    
                }
                w1 = splitSection(w1, filepath);
                for(int j = 0;j<Llen;j++){
    
                    g = 4;
                    printonE2(y, w1, array[3][j], g, filepath);

                    w1++;
                }
    
            
           
    
        return w1;
    
        }
        public static int  printP2p2(String[] array, int length, int x, String filepath, int w1){

                
                int g = 0;
              
                    splitSection(w1, filepath);
                    for(int j=0;j<length;j++){
    
                        printonE2(x, w1, array[j], g, filepath);
    
                        w1++;
                    }
               
                
    
              return w1;
    
            
        }
        public static int print1stLine(int w1, String filepath){
            try{
                
                File file = new File("C:\\SeniorScheduling\\"+filepath+".xlsx");
                FileInputStream fO = new FileInputStream(file); 
                XSSFWorkbook wb = new XSSFWorkbook(fO);
                XSSFSheet sheet = wb.getSheetAt(0);

                XSSFRow row = sheet.createRow((short)w1);

                row.createCell(0).setCellValue("Student Name");
                row.createCell(1).setCellValue("Period 1");
                row.createCell(2).setCellValue("Period 2");
                row.createCell(3).setCellValue("Period 3");
                row.createCell(4).setCellValue("Period 4");
                row.createCell(5).setCellValue("Period 5");
                row.createCell(6).setCellValue("Period 6");

                FileOutputStream f0 = new FileOutputStream("C:\\SeniorScheduling\\"+filepath+".xlsx");
                 wb.write(f0);

                 wb.close();


                 w1++;

                 return w1;

            }catch(Exception e){
                e.printStackTrace();
            }
            w1 = 2;
            return w1;
                


        }
        public static int splitSection(int w1, String filepath){
            try{
                
                File file = new File("C:\\SeniorScheduling\\"+filepath+".xlsx");
                FileInputStream fO = new FileInputStream(file); 
                XSSFWorkbook wb = new XSSFWorkbook(fO);
                XSSFSheet sheet = wb.getSheetAt(0);

                XSSFRow row = sheet.createRow((short)w1);

                row.createCell(0).setCellValue("");
               

                FileOutputStream f0 = new FileOutputStream("C:\\SeniorScheduling\\"+filepath+".xlsx");
                 wb.write(f0);

                 wb.close();


                 w1++;

                 return w1;

            }catch(Exception e){
                e.printStackTrace();
            }
            w1 = 2;
            return w1;

        }
        public static void printonE2(int schedule, int r, String name, int language, String filepath){

            try{
                
                File file = new File("C:\\SeniorScheduling\\"+filepath+".xlsx");
                FileInputStream fO = new FileInputStream(file); 
                XSSFWorkbook wb = new XSSFWorkbook(fO);
                XSSFSheet sheet = wb.getSheetAt(0);

                
                switch(schedule){

                    case 1: 

                        switch(language){
                            case 1:

                            XSSFRow row = sheet.createRow((short)r);

                            row.createCell(0).setCellValue(name);
                            row.createCell(1).setCellValue("Art");
                            row.createCell(2).setCellValue("Humane Letters");
                            row.createCell(3).setCellValue("Humane Letters");
                            row.createCell(4).setCellValue("Languagae - Spanish IV");
                            row.createCell(5).setCellValue("Physics II");
                            row.createCell(6).setCellValue("Calculus II");
                            break;

                            case 2:

                            row = sheet.createRow((short)r);

                            row.createCell(0).setCellValue(name);
                            row.createCell(1).setCellValue("Art");
                            row.createCell(2).setCellValue("Humane Letters");
                            row.createCell(3).setCellValue("Humane Letters");
                            row.createCell(4).setCellValue("Languagae - French IV");
                            row.createCell(5).setCellValue("Physics II");
                            row.createCell(6).setCellValue("Calculus II");
                            break;

                            case 3:

                            row = sheet.createRow((short)r);

                            row.createCell(0).setCellValue(name);
                            row.createCell(1).setCellValue("Art");
                            row.createCell(2).setCellValue("Humane Letters");
                            row.createCell(3).setCellValue("Humane Letters");
                            row.createCell(4).setCellValue("Languagae - German IV");
                            row.createCell(5).setCellValue("Physics II");
                            row.createCell(6).setCellValue("Calculus II");
                            break;

                            case 4:

                            row = sheet.createRow((short)r);

                            row.createCell(0).setCellValue(name);
                            row.createCell(1).setCellValue("Art");
                            row.createCell(2).setCellValue("Humane Letters");
                            row.createCell(3).setCellValue("Humane Letters");
                            row.createCell(4).setCellValue("Languagae - Greek II");
                            row.createCell(5).setCellValue("Physics II");
                            row.createCell(6).setCellValue("Calculus II");
                            break;
                        }
                    
                    break;

                    
                    case 2:
                                
                    switch(language){

                        case 1:

                        XSSFRow row = sheet.createRow((short)r);

                        row.createCell(0).setCellValue(name);
                        row.createCell(1).setCellValue("Art");
                        row.createCell(2).setCellValue("American Rhetorical Tradition");
                        row.createCell(3).setCellValue("Humane Letters");
                        row.createCell(4).setCellValue("Humane Letters");
                        row.createCell(5).setCellValue("Spanish IV");
                        row.createCell(6).setCellValue("Logic and Coding");
                        break;
                        case 2:

                        row = sheet.createRow((short)r);

                        row.createCell(0).setCellValue(name);
                        row.createCell(1).setCellValue("Art");
                        row.createCell(2).setCellValue("Humane Letters");
                        row.createCell(3).setCellValue("Humane Letters");
                        row.createCell(4).setCellValue("Languagae - French IV");
                        row.createCell(5).setCellValue("American Rhetorical Tradition");
                        row.createCell(6).setCellValue("Logic and Coding");
                        break;
                        case 3:

                        row = sheet.createRow((short)r);

                        row.createCell(0).setCellValue(name);
                        row.createCell(1).setCellValue("Art");
                        row.createCell(2).setCellValue("Humane Letters");
                        row.createCell(3).setCellValue("Humane Letters");
                        row.createCell(4).setCellValue("Languagae - German IV");
                        row.createCell(5).setCellValue("American Rhetorical Tradition");
                        row.createCell(6).setCellValue("Logic and Coding");
                        break;
                        case 4:

                        row = sheet.createRow((short)r);

                        row.createCell(0).setCellValue(name);
                        row.createCell(1).setCellValue("Art");
                        row.createCell(2).setCellValue("Humane Letters");
                        row.createCell(3).setCellValue("Humane Letters");
                        row.createCell(4).setCellValue("Languagae - Greek II");
                        row.createCell(5).setCellValue("American Rhetorical Tradition");
                        row.createCell(6).setCellValue("Logic and Coding");
                        break;
                    }
                    break;

                    case 3:

                    switch(language){

                        case 1:

                        XSSFRow row = sheet.createRow((short)r);

                        row.createCell(0).setCellValue(name);
                        row.createCell(1).setCellValue("Art");
                        row.createCell(2).setCellValue("American Rhetorical Tradition");
                        row.createCell(3).setCellValue("Humane Letters");
                        row.createCell(4).setCellValue("Humane Letters");
                        row.createCell(5).setCellValue("Language - Spanish IV");
                        row.createCell(6).setCellValue("Calculus II");
                        break;
                        case 2:

                        row = sheet.createRow((short)r);

                        row.createCell(0).setCellValue(name);
                        row.createCell(1).setCellValue("Art");
                        row.createCell(2).setCellValue("American Rhetorical Tradition");
                        row.createCell(3).setCellValue("Calculus II");
                        row.createCell(4).setCellValue("Languagae - French IV");
                        row.createCell(5).setCellValue("Humane Letters");
                        row.createCell(6).setCellValue("Humane Letters");
                        break;
                        case 3:

                        row = sheet.createRow((short)r);

                        row.createCell(0).setCellValue(name);
                        row.createCell(1).setCellValue("Art");
                        row.createCell(2).setCellValue("American Rhetorical Tradition");
                        row.createCell(3).setCellValue("Calculus II");
                        row.createCell(4).setCellValue("Languagae - German IV");
                        row.createCell(5).setCellValue("Humane Letters");
                        row.createCell(6).setCellValue("Humane Letters");
                        break;
                        case 4:

                        row = sheet.createRow((short)r);

                        row.createCell(0).setCellValue(name);
                        row.createCell(1).setCellValue("Art");
                        row.createCell(2).setCellValue("American Rhetorical Tradition");
                        row.createCell(3).setCellValue("Calculus II");
                        row.createCell(4).setCellValue("Languagae - Greek II");
                        row.createCell(5).setCellValue("Humane Letters");
                        row.createCell(6).setCellValue("Humane Letters");
                        break;
                    }
                    break;

                    case 4:

                    switch(language){

                        case 1:

                        XSSFRow row = sheet.createRow((short)r);

                        row.createCell(0).setCellValue(name);
                        row.createCell(1).setCellValue("Art");
                        row.createCell(2).setCellValue("American Rhetorical Tradition");
                        row.createCell(3).setCellValue("Physics II");
                        row.createCell(4).setCellValue("Language - Spanish IV");
                        row.createCell(5).setCellValue("Humane Letters");
                        row.createCell(6).setCellValue("Humane Letters");
                        break;
                        case 2:

                        row = sheet.createRow((short)r);

                        row.createCell(0).setCellValue(name);
                        row.createCell(1).setCellValue("Art");
                        row.createCell(2).setCellValue("American Rhetorical Tradition");
                        row.createCell(3).setCellValue("Physics II");
                        row.createCell(4).setCellValue("Languagae - French IV");
                        row.createCell(5).setCellValue("Humane Letters");
                        row.createCell(6).setCellValue("Humane Letters");
                        break;
                        case 3:

                        row = sheet.createRow((short)r);

                        row.createCell(0).setCellValue(name);
                        row.createCell(1).setCellValue("Art");
                        row.createCell(2).setCellValue("American Rhetorical Tradition");
                        row.createCell(3).setCellValue("Physics II");
                        row.createCell(4).setCellValue("Languagae - German IV");
                        row.createCell(5).setCellValue("Humane Letters");
                        row.createCell(6).setCellValue("Humane Letters");
                        break;
                        case 4:

                        row = sheet.createRow((short)r);

                        row.createCell(0).setCellValue(name);
                        row.createCell(1).setCellValue("Art");
                        row.createCell(2).setCellValue("American Rhetorical Tradition");
                        row.createCell(3).setCellValue("Physics II");
                        row.createCell(4).setCellValue("Languagae - Greek II");
                        row.createCell(5).setCellValue("Humane Letters");
                        row.createCell(6).setCellValue("Humane Letters");
                        break;
                    }
                    break;

                    case 5:

                    switch(language){

                        case 1:

                        XSSFRow row = sheet.createRow((short)r);
                        row.createCell(0).setCellValue(name);
                        row.createCell(1).setCellValue("Logic and Coding");
                        row.createCell(2).setCellValue("Art");
                        row.createCell(3).setCellValue("Humane Letters");
                        row.createCell(4).setCellValue("Humane Letters");
                        row.createCell(5).setCellValue("Language - Spanish IV");
                        row.createCell(6).setCellValue("Calculus II");
                        break;
                        case 2:

                        row = sheet.createRow((short)r);

                        row.createCell(0).setCellValue(name);
                        row.createCell(1).setCellValue("Logic and Coding");
                        row.createCell(2).setCellValue("Art");
                        row.createCell(3).setCellValue("Physics II");
                        row.createCell(4).setCellValue("Languagae - French IV");
                        row.createCell(5).setCellValue("Humane Letters");
                        row.createCell(6).setCellValue("Humane Letters");
                        break;
                        case 3:

                        row = sheet.createRow((short)r);

                        row.createCell(0).setCellValue(name);
                        row.createCell(1).setCellValue("Logic and Coding");
                        row.createCell(2).setCellValue("Art");
                        row.createCell(3).setCellValue("Physics II");
                        row.createCell(4).setCellValue("Languagae - German IV");
                        row.createCell(5).setCellValue("Humane Letters");
                        row.createCell(6).setCellValue("Humane Letters");
                        break;
                        case 4:

                        row = sheet.createRow((short)r);

                        row.createCell(0).setCellValue(name);
                        row.createCell(1).setCellValue("Logic and Coding");
                        row.createCell(2).setCellValue("Art");
                        row.createCell(3).setCellValue("Physics II");
                        row.createCell(4).setCellValue("Languagae - Greek II");
                        row.createCell(5).setCellValue("Humane Letters");
                        row.createCell(6).setCellValue("Humane Letters");
                        break;
                    }
                    break;

                    case 6:
                        switch(language){

                            case 1:

                            XSSFRow row = sheet.createRow((short)r);

                            row.createCell(0).setCellValue(name);
                            row.createCell(1).setCellValue("Art");
                            row.createCell(2).setCellValue("Humane Letters");
                            row.createCell(3).setCellValue("Humane Letters");
                            row.createCell(4).setCellValue("Language - Spanish IV");
                            row.createCell(5).setCellValue("Physics II");
                            row.createCell(6).setCellValue("Logic and Coding");
                            break;
                            case 2:

                            row = sheet.createRow((short)r);

                            row.createCell(0).setCellValue(name);
                            row.createCell(1).setCellValue("Logic and Coding");
                            row.createCell(2).setCellValue("Art");
                            row.createCell(3).setCellValue("Physics II");
                            row.createCell(4).setCellValue("Languagae - French IV");
                            row.createCell(5).setCellValue("Humane Letters");
                            row.createCell(6).setCellValue("Humane Letters");
                            break;
                            case 3:

                            row = sheet.createRow((short)r);

                            row.createCell(0).setCellValue(name);
                            row.createCell(1).setCellValue("Logic and Coding");
                            row.createCell(2).setCellValue("Art");
                            row.createCell(3).setCellValue("Physics II");
                            row.createCell(4).setCellValue("Languagae - German IV");
                            row.createCell(5).setCellValue("Humane Letters");
                            row.createCell(6).setCellValue("Humane Letters");
                            break;
                            case 4:

                            row = sheet.createRow((short)r);

                            row.createCell(0).setCellValue(name);
                            row.createCell(1).setCellValue("Logic and Coding");
                            row.createCell(2).setCellValue("Art");
                            row.createCell(3).setCellValue("Physics II");
                            row.createCell(4).setCellValue("Languagae - Greek II");
                            row.createCell(5).setCellValue("Humane Letters");
                            row.createCell(6).setCellValue("Humane Letters");
                            break;
                       
                       
                        }
                        break;
                    
                    case 7: 
                        XSSFRow row = sheet.createRow((short)r);

                        row.createCell(0).setCellValue(name);
                        row.createCell(1).setCellValue("Art");
                        row.createCell(2).setCellValue("American Rhetorical Tradition");
                        row.createCell(3).setCellValue("Humane Letters");
                        row.createCell(4).setCellValue("Humane Letters");
                        row.createCell(5).setCellValue("Physics II");
                        row.createCell(6).setCellValue("Calculus II");
                        break;
                    
                    case 8:

                        row = sheet.createRow((short)r);

                        row.createCell(0).setCellValue(name);
                            row.createCell(1).setCellValue("Logic and Coding");
                            row.createCell(2).setCellValue("Art");
                            row.createCell(3).setCellValue("Humane Letters");
                            row.createCell(4).setCellValue("Humane Letters");
                            row.createCell(5).setCellValue("Physics II");
                            row.createCell(6).setCellValue("Calculus II");
                            break;

                    case 9:

                    row = sheet.createRow((short)r);

                    row.createCell(0).setCellValue(name);
                    row.createCell(1).setCellValue("Art");
                    row.createCell(2).setCellValue("American Rhetorical Tradition");
                    row.createCell(3).setCellValue("Humane Letters");
                    row.createCell(4).setCellValue("Humane Letters");
                    row.createCell(5).setCellValue("Physics II");
                    row.createCell(6).setCellValue("Calculus II");
                    break;

                    case 10:
                    row = sheet.createRow((short)r);

                    row.createCell(0).setCellValue(name);
                    row.createCell(1).setCellValue("Logic and Coding");
                    row.createCell(2).setCellValue("Art");
                    row.createCell(3).setCellValue("Humane Letters");
                    row.createCell(4).setCellValue("Humane Letters");
                    row.createCell(5).setCellValue("American Rhetorical Tradition");
                    row.createCell(6).setCellValue("Calculus II");
                    break;
                }   

                FileOutputStream f0 = new FileOutputStream("C:\\SeniorScheduling\\"+filepath+".xlsx");
                 wb.write(f0);

                 wb.close();


            }catch(Exception e){
                e.printStackTrace();
            }

        }
        }