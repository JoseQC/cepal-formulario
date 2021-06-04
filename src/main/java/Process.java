import lombok.Getter;
import org.apache.poi.xslf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;
import java.util.stream.Stream;

@Getter
public class Process {

    public  List<String> listaVariables = new ArrayList<>();



    public void menu() throws IOException {
        System.out.println("Generate data to design forms");
        System.out.println("================================");
        Scanner myObj = new Scanner(System.in);  // Create a Scanner object
        System.out.println("Imput code process:");
        String process = myObj.nextLine();  // Read user input
        System.out.println("Process is: " + process);
        listaVariables.add(process.toUpperCase());
        //files(process);
        System.out.println("================================");
        System.out.println("Eventos y Reutilizables");
        System.out.println("================================");
        extractReuAndEven(process);
        System.out.println("=========================================================================");
        System.out.println("Excel");
        System.out.println("=========================================================================");
        listaVariables.forEach(System.out::println);
        Excel filerVariable = new Excel(listaVariables);
        filerVariable.readVarEvi();
        filerVariable.readFileInfVar();
        filerVariable.createDataTotal();
        filerVariable.createDataFilter();
        filerVariable.createDataFilterExc();
        filerVariable.getDuplicates();


    }

    private void extractReuAndEven(String pptx) throws IOException {

        boolean existReu=false;
        boolean existEve=false;
        String partPPTX = "C:\\Users\\Jose\\Workspace - Developer\\Java\\7.2.Hacienda\\"+pptx+"\\"+pptx+".pptx";

        File file = new File(partPPTX);
        int countGroup=0;
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(file));



        // get slides
        for (XSLFSlide slide : ppt.getSlides()) {
            // System.out.println("Slide: "+slide.getSlideNumber());
            for (XSLFShape shape : slide.getShapes()) {
                if(shape instanceof XSLFGroupShape){
                    XSLFGroupShape groupShape = (XSLFGroupShape)shape;
                    for (XSLFShape gShape : groupShape.getShapes()){
                        if(gShape instanceof XSLFTextShape){
                            XSLFTextShape textGShape = (XSLFTextShape) gShape;
                            if( textGShape.getText().contains("EVE_")){
                                existEve=true;
                                String codeEve=splitName(textGShape.getText());
                                System.out.println(codeEve);
                                findEvento(codeEve);
                                listaVariables.add(codeEve);
                                System.out.println(listaVariables.size());
                                System.out.println();
                            }
                            if (textGShape.getText().contains("REU_")){
                                existReu=true;
                                String codeReu=splitName(textGShape.getText());
                                System.out.println(codeReu);
                                findReutilizable(codeReu);
                                listaVariables.add(codeReu);
                                System.out.println();

                            }
                        }
                    }


                    countGroup++;
                }
            }
        }
        if(existEve)
            listaVariables.add("EVE_NOT");
        if (existReu)
            listaVariables.add("REU_NOT");


    }



    static  void findCodeEvento(String code){
        File file1 = null;
        String[] paths;

        try {

            // create new file object
            file1 = new File("C:\\Users\\Jose\\Workspace - Developer\\Java\\7.12.Eventos\\"+code);

            // array of files and directory
            paths = file1.list();

            String text="";
            // for each name in the path array
            for(String path:paths) {
                text=text+(path.substring(3, 14))+", ";
            }
            System.out.println(text.substring(0,text.length()-2));
        } catch (Exception e) {
            // if any error occurs
            e.printStackTrace();
        }
    }

    static void findEvento(String code)
    {
        File file1 = null;
        String[] paths;

        try {

            // create new file object
            file1 = new File("C:\\Users\\Jose\\Workspace - Developer\\Java\\7.12.Eventos");

            // array of files and directory
            paths = file1.list();

            // for each name in the path array
            for(String path:paths) {
                if(path.contains(code.toString())){
                    findCodeEvento(path);
                }// prints filename and directory name
            }
        } catch (Exception e) {
            // if any error occurs
            e.printStackTrace();
        }

    }

    static  void findCodeReutilizable(String code){
        File file1 = null;
        String[] paths;

        try {

            // create new file object
            file1 = new File("C:\\Users\\Jose\\Workspace - Developer\\Java\\7.11.Módulos reutilizables\\"+code);

            // array of files and directory
            paths = file1.list();
            String text="";
            // for each name in the path array
            for(String path:paths) {
                text=text+(path.substring(3, 14))+", ";
            }
            System.out.println(text.substring(0,text.length()-2));
        } catch (Exception e) {
            // if any error occurs
            e.printStackTrace();
        }
    }

    static void findReutilizable(String code)
    {
        File file1 = null;
        String[] paths;

        try {

            // create new file object
            file1 = new File("C:\\Users\\Jose\\Workspace - Developer\\Java\\7.11.Módulos reutilizables");

            // array of files and directory
            paths = file1.list();

            // for each name in the path array
            for(String path:paths) {
                if(path.contains(code.toString())){
                    findCodeReutilizable(path);
                }// prints filename and directory name
            }
        } catch (Exception e) {
            // if any error occurs
            e.printStackTrace();
        }

    }

    static String splitName(String fileName){
        String str = fileName;
        String[] arrOfStr = str.split("\\.");
        return arrOfStr[0];
    }
/*
    static  void files(String process){
        File file1 = null;
        String[] paths;
        try {
            // create new file object
            file1 = new File("C:\\Users\\Jose\\Workspace - Developer\\Java\\7.2.Hacienda");

            // array of files and directory
            paths = file1.list();

            // for each name in the path array
            for(String path:paths) {
                
                    System.out.println("files: "+path);
                // prints filename and directory name
            }
        } catch (Exception e) {
            // if any error occurs
            e.printStackTrace();
        }
    }
*/



}
