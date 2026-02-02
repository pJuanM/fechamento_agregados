# Fechamento de Agregados
Realizar análise do fechamento dos agregados da VDC LOG de forma automática

O programa consiste em algumas etapas simples:
    - Validar mês de competência = **Mês trabalhado**;
    - Verificar caso tenha algum anexo na fatura;
    - Validar forma de pagamento = **PIX**;
    - Informar conta bancária;
    - Auditar Quinzena trabalhada;
    - Verificar acordo de prestação (diária, quinzena, mensal);
    - Auditar valor de manifesto (quando diária);
    - Auditar valor de fatura (quando quinzena ou mensal);
    - Validar data de saída + finalização de manifestos (evitar ter algo em aberto);

Pós essa análise é escrita no campo de ***DESCRIÇÃO*** da própria contra fatura, afim de observar o decorrer do sancionamento de faturas
