# Verificação técnica dos dados

Esta pasta reúne os componentes utilizados na verificação técnica dos dados empregados na Prova de Conceito.

O procedimento foi estruturado em três controles complementares:

1. confronto censitário entre o histórico apresentado no SEI e os registros persistidos na tabela `historico`;
2. verificação automatizada e censitária das estruturas derivadas `vw_cessao` e `vw_cessao_fluxo` em relação aos registros de origem;
3. conferência manual complementar orientada por classes de comportamento previamente definidas a partir das regras funcionais das transformações.

A rotina computacional verifica propriedades de rastreabilidade, integridade e coerência entre os registros de origem e as estruturas derivadas, evitando, tanto quanto possível, a simples reprodução dos algoritmos utilizados na construção das views.

As planilhas resultantes da aplicação dos controles são disponibilizadas nesta pasta como evidências da execução do procedimento descrito na dissertação.

As referências ISO/IEC 25012, ISO/IEC 25024 e ISO/IEC/IEEE 29119-4 foram utilizadas como apoio conceitual à estruturação dos controles. O procedimento desenvolvido não constitui aplicação integral dessas normas nem implica declaração de conformidade com seus requisitos.
