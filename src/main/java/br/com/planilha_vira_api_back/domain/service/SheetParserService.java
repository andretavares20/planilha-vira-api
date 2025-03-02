package br.com.planilha_vira_api_back.domain.service;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import com.fasterxml.jackson.databind.ObjectMapper;

import okhttp3.MediaType;
import okhttp3.OkHttpClient;
import okhttp3.Request;
import okhttp3.RequestBody;
import okhttp3.Response;

@Service
public class SheetParserService {
    private static final Logger logger = LoggerFactory.getLogger(SheetParserService.class);

    public Map<String, Object> analisarPlanilha(InputStream inputStream) {
        logger.info("Starting to parse spreadsheet");
        Map<String, Object> rawData = new HashMap<>();
        try (OPCPackage pkg = OPCPackage.open(inputStream)) {
            logger.debug("OPCPackage opened successfully");
            XSSFWorkbook workbook = new XSSFWorkbook(pkg);
            workbook.setMissingCellPolicy(Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                if (sheet == null) {
                    logger.warn("Sheet at index {} is null, skipping", i);
                    continue;
                }
                logger.info("Extracting raw data from sheet: {} (index: {})", sheet.getSheetName(), i);
                List<Map<String, String>> sheetData = extrairDadosBrutosPlanilha(sheet);
                rawData.put(normalizarTexto(sheet.getSheetName()), sheetData);
            }
            if (rawData.isEmpty()) {
                logger.error("No valid data found in any sheet");
                throw new IllegalArgumentException("No valid data found in the spreadsheet");
            }

            // Enviar para o agente de IA externo (simulado por agora)
            Map<String, Object> structuredData = estruturarDadosIaExterna(rawData);
            logger.info("Spreadsheet parsing completed with sheets: {}", structuredData.keySet());
            return structuredData;
        } catch (Exception e) {
            logger.error("Failed to parse spreadsheet", e);
            String errorMessage = e.getMessage() != null ? e.getMessage() : "Unknown error occurred";
            throw new IllegalArgumentException("Failed to parse the spreadsheet: " + errorMessage, e);
        }
    }

    private Map<String, Object> estruturarDadosIaExterna(Map<String, Object> rawData) throws IOException {
        logger.info("Sending raw data to external AI for structuring");
    
        // Converte rawData para JSON string
        ObjectMapper mapper = new ObjectMapper();
        String rawDataJson = mapper.writeValueAsString(rawData);
    
        // Configura a chamada à API (exemplo com xAI API)
        OkHttpClient client = new OkHttpClient();
        String apiKey = "SUA_CHAVE_API_AQUI"; // Substitua pela sua chave
        String apiUrl = "https://api.xai.com/v1/completions"; // URL fictícia, ajuste conforme a API real
    
        String prompt = "Estruture os seguintes dados de uma planilha em um JSON organizado com seções e tabelas:\n" + rawDataJson;
        RequestBody body = RequestBody.create(
                MediaType.parse("application/json"),
                "{\"prompt\": \"" + prompt + "\", \"model\": \"grok\"}");
        Request request = new Request.Builder()
                .url(apiUrl)
                .post(body)
                .addHeader("Authorization", "Bearer " + apiKey)
                .addHeader("Content-Type", "application/json")
                .build();
    
        try (Response response = client.newCall(request).execute()) {
            if (!response.isSuccessful()) {
                logger.error("AI API request failed: {}", response);
                throw new IOException("Failed to get response from AI: " + response);
            }
            String responseBody = response.body().string();
            @SuppressWarnings("unchecked")
            Map<String, Object> structuredData = mapper.readValue(responseBody, Map.class);
            logger.debug("Structured data from AI: {}", structuredData);
    
            // Ajuste para acessar a estrutura da resposta da API
            Object choicesObj = structuredData.get("choices");
            if (choicesObj instanceof List) {
                @SuppressWarnings("unchecked")
                List<Map<String, Object>> choices = (List<Map<String, Object>>) choicesObj;
                if (!choices.isEmpty()) {
                    Object textObj = choices.get(0).get("text");
                    if (textObj instanceof String) {
                        // Converte o texto retornado pela IA (JSON como string) de volta para Map
                        return mapper.readValue((String) textObj, Map.class);
                    } else {
                        logger.error("AI response 'text' is not a string: {}", textObj);
                        throw new IOException("Invalid AI response format: 'text' is not a string");
                    }
                } else {
                    logger.error("AI response 'choices' is empty");
                    throw new IOException("Invalid AI response format: 'choices' is empty");
                }
            } else {
                logger.error("AI response 'choices' is not a list: {}", choicesObj);
                throw new IOException("Invalid AI response format: 'choices' is not a list");
            }
        }
    }

    private List<Map<String, String>> extrairDadosBrutosPlanilha(Sheet sheet) {
        List<Map<String, String>> sheetData = new ArrayList<>();
        int rowNum = 0;

        for (Row row : sheet) {
            if (isLinhaVazia(row))
                continue;

            Map<String, String> rowData = new HashMap<>();
            for (int i = 0; i < row.getLastCellNum(); i++) {
                String value = obterValorStringCelula(row.getCell(i));
                if (value != null && !value.trim().isEmpty()) {
                    rowData.put("col_" + i, value);
                }
            }
            if (!rowData.isEmpty()) {
                rowData.put("row_num", String.valueOf(rowNum++));
                sheetData.add(rowData);
                logger.debug("Extracted raw row: {}", rowData);
            }
        }

        logger.info("Extracted {} rows from sheet {}", sheetData.size(), sheet.getSheetName());
        return sheetData;
    }

    private boolean isLinhaVazia(Row row) {
        if (row == null)
            return true;
        for (int i = 0; i < row.getLastCellNum(); i++) {
            if (obterValorStringCelula(row.getCell(i)) != null)
                return false;
        }
        return true;
    }

    private String obterValorStringCelula(Cell cell) {
        if (cell == null)
            return null;
        try {
            switch (cell.getCellType()) {
                case STRING:
                    return cell.getStringCellValue().trim();
                case NUMERIC:
                    return String.valueOf(cell.getNumericCellValue());
                case BOOLEAN:
                    return String.valueOf(cell.getBooleanCellValue());
                default:
                    return null;
            }
        } catch (IllegalStateException e) {
            logger.warn("Error reading cell at row {} column {}: {}",
                    cell.getRowIndex(), cell.getColumnIndex(), e.getMessage());
            return null;
        }
    }

    private String normalizarTexto(String text) {
        if (text == null)
            return "";
        return text.toLowerCase()
                .replaceAll("[^a-z0-9]", "_")
                .replaceAll("_+", "_")
                .trim();
    }
}
