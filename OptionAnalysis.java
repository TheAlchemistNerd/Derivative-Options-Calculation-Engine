import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.commons.math3.distribution.NormalDistribution;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

public class OptionAnalysis {

    public static List<Map<String, Double>> loadExcelData(String filePath, String sheetName)
            throws IOException {
        List<Map<String, Double>> data = new ArrayList<>();
        System.out.println("Loading data for sheet: " + sheetName);
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                System.out.println("Sheet not found: " + sheetName);
                return data; // Return empty list if sheet doesn't exist
            }
            Iterator<Row> rowIterator = sheet.iterator();

            // Skip the header row
            if (rowIterator.hasNext()) {
                rowIterator.next();
                System.out.println("Skipping header row in sheet: " + sheetName);
            }

            int rowCount = 0;
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                rowCount++;
                System.out.println("Processing row " + rowCount + " in sheet: " + sheetName);
                Map<String, Double> rowData = new HashMap<>();

                Cell dateCell = row.getCell(0); // Date is in the first column (index 0)
                Cell priceCell = row.getCell(1); // Price is in the second column (index 1)

                Double dateValue = null;
                if (dateCell != null) {
                    if (dateCell.getCellType() == CellType.NUMERIC) {
                        dateValue = dateCell.getNumericCellValue();
                        System.out.println("  Date (numeric): " + dateValue);
                    } else if (dateCell.getCellType() == CellType.STRING) {
                        String dateStr = dateCell.getStringCellValue().trim();
                        try {
                            dateValue = Double.parseDouble(dateStr.replace("/", ""));
                            System.out.println("  Date (string): " + dateValue + " from \"" + dateStr + "\"");
                        } catch (NumberFormatException e) {
                            System.out.println("  Warning: Could not parse date string: \"" + dateStr + "\" in row " + rowCount);
                            continue; // Skip if unusable
                        }
                    } else {
                        System.out.println("  Warning: Unrecognized date cell type in row " + rowCount);
                        continue; // Skip if unusable
                    }
                } else {
                    System.out.println("  Warning: Date cell is null in row " + rowCount);
                    continue; // Skip if unusable
                }

                Double priceValue = null;
                if (priceCell != null && priceCell.getCellType() == CellType.NUMERIC) {
                    priceValue = priceCell.getNumericCellValue();
                    System.out.println("  Price: " + priceValue);
                } else if (priceCell != null && priceCell.getCellType() == CellType.STRING) {
                    System.out.println("  Warning: Price is a string in row " + rowCount + ": \"" + priceCell.getStringCellValue() + "\"");
                    continue; // Skip if unusable
                } else {
                    System.out.println("  Warning: Price cell is null or not numeric in row " + rowCount);
                    continue; // Skip if unusable
                }

                if (dateValue != null && priceValue != null) {
                    rowData.put("Date", dateValue);
                    rowData.put("Price", priceValue);
                    data.add(rowData);
                    System.out.println("  Row data added: " + rowData);
                }
            }
            System.out.println("Loaded " + data.size() + " rows from sheet: " + sheetName);
        } catch (IOException e) {
            System.err.println("Error loading data for sheet " + sheetName + ": " + e.getMessage());
            throw e;
        }
        return data;
    }

    public static List<Double> calculateReturns(List<Map<String, Double>> data) {
        List<Double> returns = new ArrayList<>();
        for (int i = 1; i < data.size(); i++) {
            double prevPrice = data.get(i - 1).get("Price");
            double currentPrice = data.get(i).get("Price");
            returns.add(Math.log(currentPrice / prevPrice));
        }
        return returns;
    }

    public static double calculateHistoricalVolatility(List<Double> returns) {
        double mean = returns.stream().mapToDouble(Double::doubleValue).average().orElse(0.0);
        double variance = returns.stream()
                .mapToDouble(r -> Math.pow(r - mean, 2))
                .average()
                .orElse(0.0);
        return Math.sqrt(variance) * Math.sqrt(252); // Annualized
    }

    public static double calculateAverageRiskFreeRate(List<Map<String, Double>> tBillData, double startDate, double endDate) {
        return tBillData.stream()
                .filter(row -> row.get("Date") >= startDate && row.get("Date") <= endDate)
                .mapToDouble(row -> row.get("Price"))
                .average()
                .orElse(0.0) / 100;
    }

    public static Map<String, Double> calculateStrikes(double spotPrice) {
        return Map.of(
                "S", spotPrice,
                "K_110", 1.1 * spotPrice,
                "K_100", spotPrice,
                "K_95", 0.95 * spotPrice
        );
    }

    public static double blackScholesCall(double S, double K, double T, double r, double sigma) {
        NormalDistribution normal = new NormalDistribution();
        double d1 = (Math.log(S / K) + (r + sigma * sigma / 2) * T) / (sigma * Math.sqrt(T));
        double d2 = d1 - sigma * Math.sqrt(T);
        return S * normal.cumulativeProbability(d1) - K * Math.exp(-r * T) * normal.cumulativeProbability(d2);
    }

    public static double blackScholesPut(double S, double K, double T, double r, double sigma) {
        NormalDistribution normal = new NormalDistribution();
        double d1 = (Math.log(S / K) + (r + sigma * sigma / 2) * T) / (sigma * Math.sqrt(T));
        double d2 = d1 - sigma * Math.sqrt(T);
        return K * Math.exp(-r * T) * normal.cumulativeProbability(-d2) - S * normal.cumulativeProbability(-d1);
    }

    public static Map<String, List<Double>> processOptions(Map<String, Double> strikes, double T, double r, double sigma) {
        double S = strikes.get("S");
        List<Double> callPrices = List.of(
                blackScholesCall(S, strikes.get("K_110"), T, r, sigma),
                blackScholesCall(S, strikes.get("K_100"), T, r, sigma),
                blackScholesCall(S, strikes.get("K_95"), T, r, sigma)
        );
        List<Double> putPrices = List.of(
                blackScholesPut(S, strikes.get("K_110"), T, r, sigma),
                blackScholesPut(S, strikes.get("K_100"), T, r, sigma),
                blackScholesPut(S, strikes.get("K_95"), T, r, sigma)
        );
        return Map.of("Call Prices", callPrices, "Put Prices", putPrices);
    }

    public static void main(String[] args) throws IOException {
        String spyFile = "SPY_Data_2024.xlsx";
        String tbillFile = "SPY_Data_2024.xlsx";

        List<Map<String, Double>> spyDataMarJun = loadExcelData(spyFile, "Mar-Jun 2024");
        List<Map<String, Double>> spyDataJulOct = loadExcelData(spyFile, "Jul-Oct 2024");
        List<Map<String, Double>> tBillData = loadExcelData(tbillFile, "TB3MS_2024");

        List<Double> returnsMarJun = calculateReturns(spyDataMarJun);
        List<Double> returnsJulOct = calculateReturns(spyDataJulOct);

        double sigmaMarJun = calculateHistoricalVolatility(returnsMarJun);
        double sigmaJulOct = calculateHistoricalVolatility(returnsJulOct);

        double rMarJun = calculateAverageRiskFreeRate(tBillData, 20240301, 20240630);
        double rJulOct = calculateAverageRiskFreeRate(tBillData, 20240701, 20241031);

        Map<String, Double> strikesMarJun = calculateStrikes(spyDataMarJun.get(0).get("Price"));
        Map<String, Double> strikesJulOct = calculateStrikes(spyDataJulOct.get(0).get("Price"));

        double T = 3.0 / 12.0;

        Map<String, List<Double>> resultsMarJun = processOptions(strikesMarJun, T, rMarJun, sigmaMarJun);
        Map<String, List<Double>> resultsJulOct = processOptions(strikesJulOct, T, rJulOct, sigmaJulOct);

        System.out.println("March-June Results: " + resultsMarJun);
        System.out.println("July-October Results: " + resultsJulOct);
    }
}
