package com.vahan;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.util.*;

import javax.xml.parsers.DocumentBuilderFactory;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;
import org.w3c.dom.*;

public class VahanExcelProcessor {

	public static void main(String[] args) {

        String inputFile = "D:/vehicle_input.xlsx";
        String outputFile = "D:/vehicle_output.xlsx";

        try (FileInputStream fis = new FileInputStream(inputFile);
             Workbook workbook = new XSSFWorkbook(fis);
             Workbook outWorkbook = new XSSFWorkbook()) {

            // STEP 1: Get Token
            String token = getToken();
            System.out.println("Token: " + token);

            Sheet sheet = workbook.getSheetAt(0);
            Sheet outSheet = outWorkbook.createSheet("Response");

            int lastRow = sheet.getLastRowNum();

            List<Map<String, String>> allData = new ArrayList<>();
            Set<String> headers = new LinkedHashSet<>();

            // STEP 2: Process vehicles
            for (int i = 1; i <= lastRow; i++) {

                Row row = sheet.getRow(i);
                if (row == null) continue;

                Cell cell = row.getCell(0);
                if (cell == null) continue;

                cell.setCellType(CellType.STRING);
                String vehicleNo = cell.getStringCellValue().trim();

                if (vehicleNo.isEmpty()) continue;

                System.out.println("Processing: " + vehicleNo);

                Map<String, String> data = getVehicleDetails(token, vehicleNo);

                // ✅ Always include vehicle number
                data.put("vehicle_no", vehicleNo);

                allData.add(data);
                headers.addAll(data.keySet());

                Thread.sleep(500);
            }

            // STEP 3: Header (vehicle_no FIRST)
            Row headerRow = outSheet.createRow(0);
            int colIndex = 0;

            headerRow.createCell(colIndex++).setCellValue("vehicle_no");

            for (String key : headers) {
                if (!key.equals("vehicle_no")) {
                    headerRow.createCell(colIndex++).setCellValue(key);
                }
            }

            // STEP 4: Write data
            int rowIndex = 1;

            for (Map<String, String> data : allData) {

                Row outRow = outSheet.createRow(rowIndex++);
                int col = 0;

                // First column = vehicle_no
                outRow.createCell(col++).setCellValue(data.getOrDefault("vehicle_no", ""));

                for (String key : headers) {
                    if (!key.equals("vehicle_no")) {
                        outRow.createCell(col++)
                                .setCellValue(data.getOrDefault(key, ""));
                    }
                }
            }

            // Auto-size columns
            for (int i = 0; i < headers.size(); i++) {
                outSheet.autoSizeColumn(i);
            }

            // Write file
            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                outWorkbook.write(fos);
            }

            System.out.println("✅ Done! Output saved at: " + outputFile);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // ================= TOKEN API =================
    public static String getToken() throws Exception {

        String urlString = "https://api-platform.mastersindia.co/api/v1/token-auth/";
        String jsonBody = "{\"username\":\"your_email\",\"password\":\"your_password\"}";

        HttpURLConnection conn = (HttpURLConnection) new URL(urlString).openConnection();

        conn.setRequestMethod("POST");
        conn.setDoOutput(true);
        conn.setRequestProperty("Content-Type", "application/json");

        try (OutputStream os = conn.getOutputStream()) {
            os.write(jsonBody.getBytes(StandardCharsets.UTF_8));
        }

        BufferedReader br = new BufferedReader(
                new InputStreamReader(conn.getInputStream(), StandardCharsets.UTF_8));

        StringBuilder response = new StringBuilder();
        String line;

        while ((line = br.readLine()) != null) {
            response.append(line);
        }

        JSONObject json = new JSONObject(response.toString());

        conn.disconnect();

        return json.getString("token");
    }

    // ================= VEHICLE API =================
    public static Map<String, String> getVehicleDetails(String token, String vehicleNo) {

        Map<String, String> result = new HashMap<>();

        try {
            String urlString = "https://api-platform.mastersindia.co/api/v2/sbt/VAHAN/";
            String jsonBody = "{\"vehiclenumber\":\"" + vehicleNo + "\"}";

            HttpURLConnection conn = (HttpURLConnection) new URL(urlString).openConnection();

            conn.setRequestMethod("POST");
            conn.setDoOutput(true);

            conn.setRequestProperty("Authorization", "JWT " + token);
            conn.setRequestProperty("Content-Type", "application/json");
            conn.setRequestProperty("Subid", "273325");
            conn.setRequestProperty("Productid", "arap");
            conn.setRequestProperty("Mode", "Buyer");

            try (OutputStream os = conn.getOutputStream()) {
                os.write(jsonBody.getBytes(StandardCharsets.UTF_8));
            }

            BufferedReader br = new BufferedReader(
                    new InputStreamReader(conn.getInputStream(), StandardCharsets.UTF_8));

            StringBuilder response = new StringBuilder();
            String line;

            while ((line = br.readLine()) != null) {
                response.append(line);
            }

            conn.disconnect();

            JSONObject json = new JSONObject(response.toString());

            JSONObject dataObj = json.getJSONObject("data");
            JSONArray respArr = dataObj.getJSONArray("response");
            JSONObject first = respArr.getJSONObject(0);

            String xml = first.getString("response");

            if (xml != null && xml.startsWith("<")) {
                result = parseXmlToMap(xml);
            }

        } catch (Exception e) {
            result.put("ERROR", e.getMessage());
        }

        return result;
    }

    // ================= XML PARSER =================
    public static Map<String, String> parseXmlToMap(String xml) {

        Map<String, String> map = new HashMap<>();

        try {
            Document doc = DocumentBuilderFactory.newInstance()
                    .newDocumentBuilder()
                    .parse(new ByteArrayInputStream(xml.getBytes()));

            doc.getDocumentElement().normalize();

            NodeList nodeList = doc.getDocumentElement().getChildNodes();

            for (int i = 0; i < nodeList.getLength(); i++) {

                Node node = nodeList.item(i);

                if (node.getNodeType() == Node.ELEMENT_NODE) {
                    map.put(node.getNodeName(), node.getTextContent());
                }
            }

        } catch (Exception e) {
            e.printStackTrace();
        }

        return map;
    }
}
