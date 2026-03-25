package com.sistema.financeiro.controller;

import com.sistema.financeiro.model.Boleto;
import com.sistema.financeiro.repository.BoletoRepository;
import com.sistema.financeiro.service.ExcelService;
import jakarta.servlet.http.HttpServletResponse;
import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.BufferedReader;
import java.io.ByteArrayInputStream;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

@Controller
@RequiredArgsConstructor
public class FinanceiroController {

    private final BoletoRepository repository;

    // --- ROTA PRINCIPAL (TELA INICIAL E FILTROS) ---
    @GetMapping("/")
    public String index(
            @RequestParam(required = false) String busca,
            @RequestParam(required = false) Integer mes,
            @RequestParam(required = false) Integer ano,
            @RequestParam(required = false) String view,
            Model model) {

        // 1. Busca todos no banco
        List<Boleto> boletos = repository.findAllByOrderByDataVencimentoDesc();

        // 2. Aplica os filtros, se o usuário tiver preenchido algo
        if (busca != null && !busca.isEmpty()) {
            boletos = boletos.stream()
                    .filter(b -> b.getEstabelecimento().toLowerCase().contains(busca.toLowerCase()))
                    .collect(Collectors.toList());
        }
        if (mes != null) {
            boletos = boletos.stream()
                    .filter(b -> b.getDataVencimento() != null && b.getDataVencimento().getMonthValue() == mes)
                    .collect(Collectors.toList());
        }
        if (ano != null) {
            boletos = boletos.stream()
                    .filter(b -> b.getDataVencimento() != null && b.getDataVencimento().getYear() == ano)
                    .collect(Collectors.toList());
        }

        // 3. Calcula o total somando os valores
        BigDecimal totalGeral = boletos.stream()
                .map(Boleto::getValor)
                .reduce(BigDecimal.ZERO, BigDecimal::add);

        // 4. Envia as variáveis para o seu HTML (Thymeleaf)
        model.addAttribute("listaBoletos", boletos);
        model.addAttribute("listaNomes", repository.findDistinctEstabelecimentos());
        model.addAttribute("totalGeral", totalGeral);

        return "index"; // Retorna o arquivo index.html
    }

    // --- SALVAR NOVO GASTO ---
    @PostMapping("/salvar")
    public String salvar(Boleto boleto) {
        repository.save(boleto);
        return "redirect:/?view=dashboard"; // Salva e volta para a aba do painel
    }

    // --- DELETAR GASTO ---
    @GetMapping("/deletar")
    public String deletar(@RequestParam Long id) {
        repository.deleteById(id);
        return "redirect:/?view=dashboard";
    }

    private final ExcelService excelService;

    // --- EXPORTAR PLANILHA CSV ---
    @GetMapping("/exportar")
    public void exportarExcel(
            @RequestParam String tipo,
            @RequestParam(required = false) Integer mes,
            @RequestParam(required = false) Integer ano,
            HttpServletResponse response) throws Exception {

        List<Boleto> boletos = repository.findAllByOrderByDataVencimentoDesc();

        // Aplica os filtros escolhidos no modal
        if ("anual".equals(tipo) && ano != null) {
            boletos = boletos.stream()
                    .filter(b -> b.getDataVencimento() != null && b.getDataVencimento().getYear() == ano)
                    .collect(Collectors.toList());
        } else if ("mensal".equals(tipo) && ano != null && mes != null) {
            boletos = boletos.stream()
                    .filter(b -> b.getDataVencimento() != null && b.getDataVencimento().getYear() == ano && b.getDataVencimento().getMonthValue() == mes)
                    .collect(Collectors.toList());
        }

        // Gera o arquivo (Os bytes) usando o nosso Service
        ByteArrayInputStream stream = excelService.gerarRelatorioExcel(boletos);

        // Configura a resposta do navegador para forçar o Download como .XLSX
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition", "attachment; filename=Relatorio_Financeiro.xlsx");

        // Escreve os bytes na saída do navegador
        org.apache.tomcat.util.http.fileupload.IOUtils.copy(stream, response.getOutputStream());
        response.flushBuffer();
    }

    // --- IMPORTAR PLANILHA CSV ---
    @PostMapping("/importar")
    public String importarCSV(@RequestParam("file") MultipartFile file) {
        if (file.isEmpty()) return "redirect:/?view=dashboard";

        try (BufferedReader reader = new BufferedReader(new InputStreamReader(file.getInputStream(), StandardCharsets.UTF_8))) {
            String line;
            boolean isFirstLine = true;
            List<Boleto> novosBoletos = new ArrayList<>();
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd/MM/yyyy");

            while ((line = reader.readLine()) != null) {
                if (isFirstLine) { isFirstLine = false; continue; } // Pula o cabeçalho

                String[] colunas = line.split(","); // Separador do CSV
                if (colunas.length >= 3) {
                    Boleto b = new Boleto();
                    try {
                        // Espera formato dd/MM/yyyy no CSV, se for diferente, pode ajustar aqui
                        b.setDataVencimento(LocalDate.parse(colunas[0].trim(), formatter));
                        b.setEstabelecimento(colunas[1].trim());
                        b.setValor(new BigDecimal(colunas[2].trim()));
                        if (colunas.length >= 4) b.setDetalhes(colunas[3].trim());

                        novosBoletos.add(b);
                    } catch (Exception e) {
                        System.out.println("Erro ao ler linha: " + line); // Ignora linhas mal formatadas
                    }
                }
            }
            repository.saveAll(novosBoletos); // Salva tudo de uma vez
        } catch (Exception e) {
            e.printStackTrace();
        }
        return "redirect:/?view=dashboard";
    }
}