CREATE TABLE controle (codigo INT IDENTITY PRIMARY KEY NOT NULL, nome_cliente VARCHAR(200), empresa_cliente VARCHAR(200), contato_cliente VARCHAR(100), valor_servico DECIMAL(14,2), descricao_servico VARCHAR(1000), status_servico INT, data_entrada DATE, data_saida DATE, data_atualizacao DATE);
SELECT * FROM controle
DROP TABLE controle