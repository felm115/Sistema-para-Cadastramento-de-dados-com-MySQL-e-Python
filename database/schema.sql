CREATE DATABASE IF NOT EXISTS sistema_alunos;
USE sistema_alunos;

CREATE TABLE IF NOT EXISTS alunos (
    id_aluno VARCHAR(20) PRIMARY KEY,
    nome VARCHAR(120) NOT NULL,
    idade INT,
    genero VARCHAR(30),
    contato VARCHAR(30),
    email VARCHAR(120),
    criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

INSERT INTO alunos (id_aluno, nome, idade, genero, contato, email)
VALUES
('ALU101', 'João Silva', 22, 'Masculino', '11999990001', 'joao.silva@email.com'),
('ALU102', 'Maria Oliveira', 24, 'Feminino', '11999990002', 'maria.oliveira@email.com'),
('ALU103', 'Carlos Souza', 21, 'Masculino', '11999990003', 'carlos.souza@email.com')
ON DUPLICATE KEY UPDATE
nome = VALUES(nome),
idade = VALUES(idade),
genero = VALUES(genero),
contato = VALUES(contato),
email = VALUES(email);