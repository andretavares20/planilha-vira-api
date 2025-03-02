package br.com.planilha_vira_api_back.domain.service;

import java.util.UUID;

import org.springframework.stereotype.Service;

@Service
public class ApiKeyGenerator {
    public String gerarChave() {
        return UUID.randomUUID().toString();
    }
}
