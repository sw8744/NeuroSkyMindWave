package com.github.jasgo.neurosky;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;

import java.io.*;
import java.net.Socket;
import java.nio.charset.StandardCharsets;

public class ThinkGearClient {
    
    private final String name;
    private final String key;
    private final String host;
    private final int port;

    public ThinkGearClient(String name, String key) {
        this.name = name;
        this.key = key;
        this.host = "210.114.22.146";
        this.port = 13854;
    }

    public String getName() {
        return name;
    }

    public String getKey() {
        return key;
    }

    public String getHost() {
        return host;
    }

    public int getPort() {
        return port;
    }
    public JSONObject getAuth() {
        JSONObject result = new JSONObject();
        result.put("appName", name);
        result.put("appKey", key);
        return result;
    }
    public JSONObject getConfig() {
        JSONObject result = new JSONObject();
        result.put("enableRawOutput", true);
        result.put("format", "Json");
        return result;
    }
    public void connect() throws IOException, ParseException {
        System.out.println(getAuth().toJSONString());
        System.out.println(getConfig().toJSONString());
        Socket socket = new Socket("localhost", port);
        BufferedReader reader = new BufferedReader(new InputStreamReader(socket.getInputStream(), StandardCharsets.UTF_8));
        PrintWriter writer = new PrintWriter(socket.getOutputStream(), true, StandardCharsets.UTF_8);
        String value;
        boolean configSent = false;
        writer.println(getAuth().toJSONString());
        
        Scanner sc = new Scanner(System.in);
        String fname = sc.nextLine();
        sc.close();

        File file = new File("C:\\" + fname + ".xlsx");
        FileOutputStream fos = new FileOutputStream(file);

        Workbook xlsxWb = new XSSFWorkbook();
        Sheet sheet1 = xlsxWb.createSheet(fname);


        Row row = null;
        Cell cell = null;
        
        row = sheet1.createRow(0);
        cell = row.createCell(0);
        cell.setCellValue("attention");
        cell = row.createCell(1);
        cell.setCellValue("meditation");
        cell = row.createCell(2);
        cell.setCellValue("delta");
        cell = row.createCell(3);
        cell.setCellValue("theta");
        cell = row.createCell(4);
        cell.setCellValue("lowAlpha");
        cell = row.createCell(5);
        cell.setCellValue("highAlpha");
        cell = row.createCell(6);
        cell.setCellValue("lowBeta");
        cell = row.createCell(7);
        cell.setCellValue("highBeta");
        cell = row.createCell(8);
        cell.setCellValue("lowGamma");
        cell = row.createCell(9);
        cell.setCellValue("highGamma");
        cell = row.createCell(10);
        cell.setCellValue("poorSignalLevel");

        int row_num = 1;
        while ((value = reader.readLine()) != null) {
            if (!configSent) {
                configSent = true;
                writer.println(getConfig().toJSONString());
            } else {
                System.out.println(value);
                
                JSONParser jsonParser = new JSONParser();
                JSONObject value_json = (JSONObject) jsonParser.parse(value);

                row = sheet1.createRow(row_num);
                JSONObject eSense = (JSONObject) value_json.get("eSense");
                JSONObject eegPower = (JSONObject) value_json.get("eegPower");
                cell = row.createCell(0);
                cell.setCellValue((String) eSense.get("attention"));
                cell = row.createCell(1);
                cell.setCellValue((String) eSense.get("meditation"));
                cell = row.createCell(2);
                cell.setCellValue((String) eegPower.get("delta"));
                cell = row.createCell(3);
                cell.setCellValue((String) eegPower.get("theta"));
                cell = row.createCell(4);
                cell.setCellValue((String) eegPower.get("lowAlpha"));
                cell = row.createCell(5);
                cell.setCellValue((String) eegPower.get("highAlpha"));
                cell = row.createCell(6);
                cell.setCellValue((String) eegPower.get("lowBeta"));
                cell = row.createCell(7);
                cell.setCellValue((String) eegPower.get("highBeta"));
                cell = row.createCell(8);
                cell.setCellValue((String) eegPower.get("lowGamma"));
                cell = row.createCell(9);
                cell.setCellValue((String) eegPower.get("highGamma"));
                cell = row.createCell(10);
                cell.setCellValue((String) value_json.get("poorSignalLevel"));
                row_num += 1;
                xlsxWb.write(fos);
                if(fos != null) {
                    fos.close();

            }
        }
    }
}
