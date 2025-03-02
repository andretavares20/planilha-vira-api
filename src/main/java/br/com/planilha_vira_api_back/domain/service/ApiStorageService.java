package br.com.planilha_vira_api_back.domain.service;

import java.util.HashMap;
import java.util.Map;
import java.util.Optional;

import org.springframework.stereotype.Service;

import br.com.planilha_vira_api_back.domain.model.GeneratedApi;

@Service
public class ApiStorageService {
    private final Map<String, GeneratedApi> storage = new HashMap<>();

    public void salvar(GeneratedApi api) {
        storage.put(api.endpointId(), api);
    }

    public Optional<Map<String, Object>> buscarDadosPorIdEChave(String endpointId, String providedKey) {
        return Optional.ofNullable(storage.get(endpointId))
                .filter(api -> api.apiKey().equals(providedKey))
                .map(GeneratedApi::data);
    }
}
