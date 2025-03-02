package br.com.planilha_vira_api_back.controller;

import java.io.IOException;
import java.util.Map;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RequestPart;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import br.com.planilha_vira_api_back.domain.model.ApiGenerationResponse;
import br.com.planilha_vira_api_back.domain.model.GeneratedApi;
import br.com.planilha_vira_api_back.domain.service.ApiKeyGenerator;
import br.com.planilha_vira_api_back.domain.service.ApiStorageService;
import br.com.planilha_vira_api_back.domain.service.SheetParserService;
import lombok.RequiredArgsConstructor;

@RestController
@RequestMapping("/api")
@RequiredArgsConstructor
public class ApiGenerationController {

    @Value("${api.base-url:http://localhost:8080/api/}")
    private String baseUrl;

    private final ApiStorageService storageService;
    private final ApiKeyGenerator keyGenerator;
    private final SheetParserService sheetParserService;

    @PostMapping(value = "/generate", consumes = "multipart/form-data")
    public ResponseEntity<ApiGenerationResponse> gerarApi(@RequestPart("file") MultipartFile file) {
        try {
            GeneratedApi api = criarApi(file);
            storageService.salvar(api);

            ApiGenerationResponse response = new ApiGenerationResponse(baseUrl + api.endpointId(), api.apiKey());
            return ResponseEntity.ok(response);
        } catch (IllegalArgumentException e) {
            return badRequest("Erro ao processar a planilha: " + e.getMessage());
        } catch (IOException e) {
            return internalServerError("Erro interno ao ler o arquivo: " + e.getMessage());
        }
    }

    @GetMapping("/{id}")
    public ResponseEntity<Map<String, Object>> obterDadosApi(@PathVariable String id, @RequestParam String key) {
        return storageService.buscarDadosPorIdEChave(id, key)
                .map(ResponseEntity::ok)
                .orElseGet(() -> ResponseEntity.status(HttpStatus.UNAUTHORIZED)
                        .body(Map.of("error", "Invalid API Key")));
    }

    private GeneratedApi criarApi(MultipartFile file) throws IOException {
        Map<String, Object> parsedData = sheetParserService.analisarPlanilha(file.getInputStream());
        String endpointId = gerarIdEndpoint();
        String apiKey = keyGenerator.gerarChave();
        return new GeneratedApi(endpointId, apiKey, parsedData);
    }

    private String gerarIdEndpoint() {
        return java.util.UUID.randomUUID().toString().substring(0, 8);
    }

    private ResponseEntity<ApiGenerationResponse> badRequest(String message) {
        return ResponseEntity.status(HttpStatus.BAD_REQUEST)
                .body(new ApiGenerationResponse(null, null, message));
    }

    private ResponseEntity<ApiGenerationResponse> internalServerError(String message) {
        return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                .body(new ApiGenerationResponse(null, null, message));
    }
}
