package com.sistema.financeiro.service;

import com.sistema.financeiro.model.Boleto;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.time.format.TextStyle;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.stream.Collectors;

@Service
public class ExcelService {

    public ByteArrayInputStream gerarRelatorioExcel(List<Boleto> boletos) {
        try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {

            // --- 1. AGRUPAMENTO POR MÊS E ANO ---
            Map<String, List<Boleto>> boletosPorMes = boletos.stream()
                    .filter(b -> b.getDataVencimento() != null)
                    .collect(Collectors.groupingBy(b -> {
                        String mes = b.getDataVencimento().getMonth().getDisplayName(TextStyle.FULL, new Locale("pt", "BR")).toUpperCase();
                        int ano = b.getDataVencimento().getYear();
                        return mes + " " + ano;
                    }));

            // --- 2. CRIAÇÃO DOS ESTILOS ---
            CellStyle headerStyle = createHeaderStyle(workbook);
            CellStyle currencyStyle = createCurrencyStyle(workbook);
            CellStyle dateStyle = createDateStyle(workbook);
            CellStyle groupStyle = createGroupStyle(workbook); // Estilo da faixa cinza do Estabelecimento
            CellStyle subtotalStyle = createSubtotalStyle(workbook); // Estilo para o Total do Estabelecimento
            CellStyle totalGeralStyle = createTotalGeralStyle(workbook); // Estilo para o Total do Mês

            if (boletosPorMes.isEmpty()) {
                workbook.createSheet("Relatorio Vazio");
            }

            // --- 3. LOOP DE ABAS (SHEETS) ---
            for (Map.Entry<String, List<Boleto>> entry : boletosPorMes.entrySet()) {
                String nomeAba = entry.getKey();
                List<Boleto> boletosDoMes = entry.getValue();

                Sheet sheet = workbook.createSheet(nomeAba);

                // Variável para controlar em qual linha estamos escrevendo
                int rowIdx = 0;

                // --- 4. CABEÇALHO DA PLANILHA (Reduzido para 3 colunas) ---
                Row headerRow = sheet.createRow(rowIdx++);
                String[] colunas = {"Vencimento", "Detalhes", "Valor"}; // "Data de Vencimento" alterado para "Vencimento"
                for (int i = 0; i < colunas.length; i++) {
                    Cell cell = headerRow.createCell(i);
                    cell.setCellValue(colunas[i]);
                    cell.setCellStyle(headerStyle);
                }

                // --- 5. AGRUPAMENTO POR ESTABELECIMENTO DENTRO DO MÊS ---
                Map<String, List<Boleto>> porEstabelecimento = boletosDoMes.stream()
                        .collect(Collectors.groupingBy(b -> b.getEstabelecimento() != null ? b.getEstabelecimento() : "Outros"));

                double totalGeralDoMes = 0;

                // Loop para imprimir cada estabelecimento
                for (Map.Entry<String, List<Boleto>> estabEntry : porEstabelecimento.entrySet()) {
                    String nomeEstab = estabEntry.getKey();
                    List<Boleto> listaDoEstab = estabEntry.getValue();

                    // 5.1 Linha de Título do Grupo (Faixa Cinza mesclando as 3 colunas)
                    Row groupRow = sheet.createRow(rowIdx++);
                    Cell cellGroup = groupRow.createCell(0);
                    cellGroup.setCellValue(nomeEstab.toUpperCase());
                    cellGroup.setCellStyle(groupStyle);
                    // Mescla da coluna 0 (Vencimento) até a 2 (Valor) na linha atual
                    sheet.addMergedRegion(new CellRangeAddress(rowIdx - 1, rowIdx - 1, 0, 2));

                    double subtotalEstab = 0;

                    // 5.2 Imprime os gastos daquele estabelecimento
                    for (Boleto b : listaDoEstab) {
                        Row row = sheet.createRow(rowIdx++);

                        // Coluna 0: Data
                        Cell dateCell = row.createCell(0);
                        dateCell.setCellValue(b.getDataVencimento());
                        dateCell.setCellStyle(dateStyle);

                        // Coluna 1: Detalhes
                        Cell detCell = row.createCell(1);
                        detCell.setCellValue(b.getDetalhes() != null && !b.getDetalhes().isEmpty() ? b.getDetalhes() : "-");

                        // Coluna 2: Valor
                        Cell valCell = row.createCell(2);
                        double valorNum = b.getValor().doubleValue();
                        valCell.setCellValue(valorNum);
                        valCell.setCellStyle(currencyStyle);

                        subtotalEstab += valorNum;
                    }

                    // 5.3 Linha de Subtotal do Estabelecimento
                    Row subtotalRow = sheet.createRow(rowIdx++);

                    // Texto do Subtotal na Coluna 1 (Detalhes) para ficar colado no valor
                    Cell labelSubtotal = subtotalRow.createCell(1);
                    labelSubtotal.setCellValue("Total " + nomeEstab + ":");

                    // Valor do Subtotal na Coluna 2 (Valor)
                    Cell valueSubtotal = subtotalRow.createCell(2);
                    valueSubtotal.setCellValue(subtotalEstab);
                    valueSubtotal.setCellStyle(subtotalStyle);

                    // Adiciona uma linha em branco para separar do próximo estabelecimento
                    rowIdx++;

                    // Soma no Total Geral
                    totalGeralDoMes += subtotalEstab;
                }

                // --- 6. LINHA DE TOTAL GERAL DO MÊS ---
                Row totalRow = sheet.createRow(rowIdx++);

                // Colocamos o texto "TOTAL GERAL:" na coluna 1 (Detalhes), alinhado à direita
                Cell labelTotalGeral = totalRow.createCell(1);
                labelTotalGeral.setCellValue("TOTAL GERAL:");
                labelTotalGeral.setCellStyle(totalGeralStyle);

                // Valor total na coluna 2 (Valor)
                Cell valueTotalGeral = totalRow.createCell(2);
                valueTotalGeral.setCellValue(totalGeralDoMes);
                valueTotalGeral.setCellStyle(totalGeralStyle); // Usa o mesmo fundo preto/texto branco

                // Para garantir que o estilo de moeda aplique no total geral também:
                CellStyle totalGeralCurrency = workbook.createCellStyle();
                totalGeralCurrency.cloneStyleFrom(totalGeralStyle);
                totalGeralCurrency.setDataFormat(workbook.createDataFormat().getFormat("R$ #,##0.00"));
                valueTotalGeral.setCellStyle(totalGeralCurrency);

                // --- 7. AJUSTE DE LARGURA DAS COLUNAS ---
                sheet.autoSizeColumn(0); // Vencimento
                sheet.setColumnWidth(1, 256 * 30); // Define largura fixa para Detalhes (30 caracteres) para não ficar esmagado
                sheet.autoSizeColumn(2); // Valor
            }

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());

        } catch (Exception e) {
            throw new RuntimeException("Falha ao exportar Excel: " + e.getMessage());
        }
    }

    // --- MÉTODOS AUXILIARES DE ESTILO ---

    private CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font font = workbook.createFont();
        font.setColor(IndexedColors.WHITE.getIndex());
        font.setBold(true);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        return style;
    }

    private CellStyle createCurrencyStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setDataFormat(workbook.createDataFormat().getFormat("R$ #,##0.00"));
        return style;
    }

    private CellStyle createDateStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setDataFormat(workbook.createDataFormat().getFormat("dd/mm/yyyy"));
        return style;
    }

    private CellStyle createGroupStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
    }

    private CellStyle createSubtotalStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setDataFormat(workbook.createDataFormat().getFormat("R$ #,##0.00"));
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
    }

    private CellStyle createTotalGeralStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.BLACK.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font font = workbook.createFont();
        font.setColor(IndexedColors.WHITE.getIndex());
        font.setBold(true);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.RIGHT); // Alinha o texto "TOTAL GERAL:" à direita
        return style;
    }
}