# Painel de Audiências

Este projeto é um painel web para visualizar audiências jurídicas, com ênfase na contagem regressiva e destaque da próxima audiência.

## Tecnologias Usadas
- HTML
- CSS
- JavaScript

## Como Usar
1. Clone o repositório:
```bash
git clone https://github.com/ErickFJSantos314/painel-audiencias.git
```
2. Abra o arquivo `index.html` no seu navegador.
3. Clique em "escolher arquivo" no canto inferior esquerdo da tela
4. Selecione um arquivo xlsx ⚠️ Este código obedece apenas o modelo de planílha estabelecido em `modelo-planilha.xlsx`
5. Você pode alternar entre audiências do dia e audiências da semana clicando nos botões "alterar exibição"
6. A audiência mais próxima sempre aparecerá no card destaque, se não houver uma próxima audiência no dia de hoje, isso será informado no card
7. Sempre que faltarem 5 minutos para a próxima audiência, o card destaque alertará piscando em amarelo
8. Entre o minuto exato de inicio da audiência e 1 minuto depois, o card piscará em vermelho
9. Passadas as audiências, elas irão para uma aba nomeada como "já realizadas", que fica embaixo da tabela
10. Existem 3 cards abaixo da tabela indicando qual o total de audiências hoje, amanhã e na semana

## Contato
Se você tiver alguma dúvida ou sugestão, entre em contato comigo pelo e-mail: erickfranciscojs@hotmail.com.
