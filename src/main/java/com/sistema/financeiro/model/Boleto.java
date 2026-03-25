package com.sistema.financeiro.model;

import jakarta.persistence.*;
import lombok.Data;
import org.springframework.format.annotation.DateTimeFormat;

import java.math.BigDecimal;
import java.time.LocalDate;

@Entity
@Data // O Lombok cria os Getters e Setters automaticamente por baixo dos panos
public class Boleto {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    @DateTimeFormat(pattern = "yyyy-MM-dd")
    private LocalDate dataVencimento;

    private String estabelecimento;

    private String detalhes;

    private BigDecimal valor;
}