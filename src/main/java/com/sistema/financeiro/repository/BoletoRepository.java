package com.sistema.financeiro.repository;

import com.sistema.financeiro.model.Boleto;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;

import java.util.List;

@Repository
public interface BoletoRepository extends JpaRepository<Boleto, Long> {

    // Retorna a lista de todos os boletos ordenados do mais recente para o mais antigo
    List<Boleto> findAllByOrderByDataVencimentoDesc();

    // Busca apenas os nomes únicos dos estabelecimentos para preencher o seu "select"
    @Query("SELECT DISTINCT b.estabelecimento FROM Boleto b WHERE b.estabelecimento IS NOT NULL ORDER BY b.estabelecimento")
    List<String> findDistinctEstabelecimentos();
}