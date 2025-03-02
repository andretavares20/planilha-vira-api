package br.com.planilha_vira_api_back.domain.model;

public record ApiGenerationResponse(String endpoint, String key, String errorMessage) {
    public ApiGenerationResponse(String endpoint, String key) {
        this(endpoint, key, null);
    }
}
