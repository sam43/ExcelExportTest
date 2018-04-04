package com.example.sayem.excelexporttest;

import android.content.Context;
import android.os.Environment;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.util.Log;
import android.view.View;
import android.widget.Button;
import android.widget.Toast;

import com.example.sayem.excelexporttest.model.Contact;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.ObjectInputStream;
import java.io.ObjectOutput;
import java.io.ObjectOutputStream;
import java.util.ArrayList;
import java.util.List;

public class ExcelActivity extends AppCompatActivity {
    private static String[] columns = {"First Name", "Last Name", "Email", "Date Of Birth"};
    private static List<Contact> contacts = new ArrayList<Contact>();
    File file;
    Button save;
    Workbook workbook;
    final static String fileName = "contacts.xlsx";
    //  final static String path = Environment.getExternalStorageDirectory().getAbsolutePath().toString() + "/instinctcoder/readwrite/" ;
    // final static String path = Environment.DIRECTORY_DOCUMENTS + "aaTutorial";
    final static String path = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOCUMENTS) + "/m";
    final static String TAG = FileHelper.class.getName();
    //private String filename = "seeker_list.xls";

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_excel);
        save = findViewById(R.id.save);
        //PopulateList();
        //ExcelDoings();
        //create blank workbook
        XSSFWorkbook workbook = new XSSFWorkbook();

        //Create a blank sheet
        XSSFSheet sheet = workbook.createSheet("Country");

        ArrayList<Object[]> data=new ArrayList<Object[]>();
        data.add(new String[]{"Country","Capital","Population"});
        data.add(new Object[]{"India","Delhi",10000});
        data.add(new Object[]{"France","Paris",40000});
        data.add(new Object[]{"Germany","Berlin",20000});
        data.add(new Object[]{"England","London",30000});


        //Iterate over data and write to sheet
        int rownum = 0;
        for (Object[] countries : data)
        {
            Row row = sheet.createRow(rownum++);

            int cellnum = 0;
            for (Object obj : countries)
            {
                Cell cell = row.createCell(cellnum++);
                if(obj instanceof String)
                    cell.setCellValue((String)obj);
                else if(obj instanceof Double)
                    cell.setCellValue((Double)obj);
                else if(obj instanceof Integer)
                    cell.setCellValue((Integer)obj);
            }
        }
        try
        {
            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File("CountriesDetails.xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("CountriesDetails.xlsx has been created successfully");
        }
        catch (IOException e)
        {
            e.printStackTrace();
        }

}

    private void ExcelDoings() {

        workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Contacts");
        Font headerFont = workbook.createFont();
        headerFont.setItalic(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.RED.getIndex());

        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);

        // Create a Row
        Row headerRow = sheet.createRow(0);

        for (int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);
        }

        // Create Other rows and cells with contacts data
        int rowNum = 1;

        for (Contact contact : contacts) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(contact.getFirstName());
            row.createCell(1).setCellValue(contact.getLastName());
            row.createCell(2).setCellValue(contact.getEmail());
            row.createCell(3).setCellValue(contact.getDateOfBirth());
        }

        // Resize all columns to fit the content size
        /*for (int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(i);
        }*/

/*
        FileOutputStream fileOut = null;
        try {
            fileOut = new FileOutputStream("contacts.xlsx");
            workbook.write(fileOut);
            fileOut.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
*/

//        sheet.autoSizeColumn(2);

        //SaveAsXslsFile();

        save.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                try
                {
                    //Write the workbook in file system
                    FileOutputStream out = new FileOutputStream(new File(fileName));
                    workbook.write(out);
                    out.close();
                    Log.d(TAG,"CountriesDetails.xlsx has been created successfully");
                }
                catch (IOException e)
                {
                    e.printStackTrace();
                }
                //Toast.makeText(ExcelActivity.this,"Saved to file"+FileHelper.path,Toast.LENGTH_SHORT).show();

            }
        });

    }

    private void SaveAsXslsFile() {

        ObjectOutput out = null;
        FileInputStream fis;
        try {
            new File(path).mkdir();

            File file = new File(path, fileName);

            if (!file.exists()) {
                file.createNewFile();
            }
            /*byte[] byteArray = new byte[file.length()]();
            byteArray = FileUtils.readFileToByteArray(file);*/

            /*FileOutputStream fileOutputStream = new FileOutputStream(file, true);
            //fileOutputStream.write((data + System.getProperty("line.separator")).getBytes());
            fileOutputStream.write(contacts);*/

            fis = this.openFileInput("CalEvents");
            ObjectInputStream ois = new ObjectInputStream(fis);
            ArrayList<Contact> returnlist = (ArrayList<Contact>) ois.readObject();
            Log.d(TAG, "SaveAsXslsFile: " + returnlist.size());
            ois.close();

        } catch (IOException ex) {
            Log.d(TAG, ex.getMessage());
        } catch (ClassNotFoundException e) {
            e.printStackTrace();
        }
    }


    private void PopulateList() {

        contacts.add(new Contact("Sylvain", "Saurel",
                "sylvain.saurel@gmail.com", "17/01/1980"));
        contacts.add(new Contact("Albert", "Dupond",
                "sylvain.saurel@gmail.com", "17/08/1989"));
        contacts.add(new Contact("Pierre", "Dupont",
                "sylvain.saurel@gmail.com", "17/07/1956"));
        contacts.add(new Contact("Mariano", "Diaz", "sylvain.saurel@gmail.com",
                "17/05/1988"));
        contacts.add(new Contact("Ahmmod", "Chapha",
                "bd.chapha@gmail.com", "17/08/1989"));
        contacts.add(new Contact("Abdur", "Razzaq",
                "abd.razz21@gmail.com", "17/07/1911"));
        contacts.add(new Contact("Assa", "Reas",
                "ss.saul@gmail.com", "17/05/1988"));

        Toast.makeText(this, "Populate data done!", Toast.LENGTH_SHORT).show();
        Log.d("LOGG", "populated");
    }

    public void saveFile(View view) {

        try
        {
            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File(path + fileName));
            workbook.write(out);
            out.close();
            System.out.println("CountriesDetails.xlsx has been created successfully");
        }
        catch (IOException e)
        {
            e.printStackTrace();
        }
            Toast.makeText(ExcelActivity.this,"Saved to file"+FileHelper.path,Toast.LENGTH_SHORT).show();
    }

    public void wrtieFileOnInternalStorage(Context mcoContext, String sFileName, String sBody){
        File file = new File(mcoContext.getFilesDir(),"mydir");
        if(!file.exists()){
            file.mkdir();
        }

        try{
            File gpxfile = new File(file, sFileName);
            FileWriter writer = new FileWriter(gpxfile);
            writer.append(sBody);
            writer.flush();
            writer.close();

        }catch (Exception e){
            e.printStackTrace();

        }
    }
}
