package com.example.sayem.excelexporttest;

import android.content.Context;
import android.os.Build;
import android.os.Environment;
import android.support.annotation.RequiresApi;
import android.util.Log;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;

@RequiresApi(api = Build.VERSION_CODES.KITKAT)
public class FileHelper {

    final static String fileName = "contacts.xlsx";
    //  final static String path = Environment.getExternalStorageDirectory().getAbsolutePath().toString() + "/instinctcoder/readwrite/" ;
    // final static String path = Environment.DIRECTORY_DOCUMENTS + "aaTutorial";
    final static String path = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOCUMENTS)+"/m";
    final static String TAG = FileHelper.class.getName();

    public static String ReadFile(Context context) {
        String line = null;

        try {
            FileInputStream fileInputStream = null;
            try {

                fileInputStream = new FileInputStream(new File(path + fileName));
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            }
            InputStreamReader inputStreamReader = new InputStreamReader(fileInputStream);
            BufferedReader bufferedReader = new BufferedReader(inputStreamReader);
            StringBuilder stringBuilder = new StringBuilder();

            while ((line = bufferedReader.readLine()) != null) {
                stringBuilder.append(line + System.getProperty("line.separator"));
            }
            fileInputStream.close();
            line = stringBuilder.toString();

            bufferedReader.close();
        } catch (java.io.IOException ex) {
            Log.d(TAG, ex.getMessage());
        }
        return line;
    }

    public static boolean saveToFile(String data) {
        try {
            new File(path).mkdir();

            File file = new File(path, fileName);

            if (!file.exists()) {
                file.createNewFile();
            }
            FileOutputStream fileOutputStream = new FileOutputStream(file, true);
            //fileOutputStream.write((data + System.getProperty("line.separator")).getBytes());
            fileOutputStream.write(data.getBytes());
            return true;
        } catch (FileNotFoundException ex) {
            Log.d(TAG, ex.getMessage());
        } catch (IOException ex) {
            Log.d(TAG, ex.getMessage());
        }
        return false;


    }
}
