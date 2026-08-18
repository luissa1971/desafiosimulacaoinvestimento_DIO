# Simulador de Investimentos em Fundos Imobiliários

Projeto desenvolvido como parte do desafio de Excel da [Digital Innovation One (DIO)](https://www.dio.me/). A proposta foi construir uma ferramenta que permita simular aportes mensais em fundos imobiliários, estimar o patrimônio acumulado e projetar os dividendos mensais.

![Tela principal do simulador](images/simulador-fii.png)

## Objetivo

Aplicar recursos do Microsoft Excel na criação de uma solução prática para responder perguntas comuns de investidores:

- Quanto investir por mês?
- Por quanto tempo realizar os aportes?
- Qual patrimônio poderá ser acumulado?
- Qual valor mensal de dividendos poderá ser obtido?
- Como distribuir o aporte entre diferentes tipos de fundos imobiliários?

## Funcionalidades

- Cálculo automático da sugestão de investimento com base em 30% do salário informado;
- Simulação de aportes mensais com juros compostos;
- Projeção do patrimônio acumulado no período escolhido;
- Estimativa de dividendos mensais;
- Comparação de cenários de 2, 5, 10, 20 e 30 anos;
- Seleção de perfil de investidor: Conservador, Moderado ou Agressivo;
- Distribuição automática do aporte entre fundos de Papel, Tijolo, Híbridos, FOFs, Desenvolvimento e Hotelarias;
- Gráfico para visualizar a composição sugerida da carteira;
- Validação de dados nas principais células de entrada.

## Estrutura da planilha

### Aba `APP`

É a tela principal do simulador. As células em azul são editáveis e os demais campos são calculados automaticamente.

Principais entradas:

- Salário;
- Rendimento mensal da carteira;
- Valor do aporte mensal;
- Prazo do investimento em anos;
- Taxa de rendimento mensal;
- Perfil do investidor.

### Aba `Planilha2`

Contém a tabela auxiliar usada para relacionar cada perfil de investidor aos percentuais sugeridos para os tipos de FII. Essa tabela alimenta a tela principal por meio da função `PROCV` (`VLOOKUP`).

## Fórmulas utilizadas

### Sugestão de investimento

```excel
=Salário*30%
```

### Patrimônio acumulado

```excel
=VF(Taxa_Mensal;Quantidade_de_Anos*12;Aporte*-1)
```

Na versão em inglês do Excel, a função aparece como `FV`.

### Dividendos mensais

```excel
=Patrimônio_Acumulado*Rendimento_da_Carteira
```

### Percentual sugerido por perfil

```excel
=PROCV(Perfil&"-"&Tipo_de_FII;Tabela_de_Perfis;4;FALSO)
```

### Valor destinado a cada tipo de FII

```excel
=Percentual_Sugerido*Aporte_Mensal
```

## Como usar

1. Baixe o arquivo `Simulador_Investimentos_FII.xlsx`;
2. Abra o arquivo no Microsoft Excel;
3. Na aba `APP`, altere somente as células em azul;
4. Informe salário, taxas, aporte, prazo e perfil;
5. Observe a atualização automática dos resultados, cenários, distribuição e gráfico.

## Tecnologias e recursos aplicados

- Microsoft Excel;
- Fórmulas financeiras;
- Juros compostos;
- Função `VF`/`FV`;
- Função `PROCV`/`VLOOKUP`;
- Nomes definidos;
- Validação de dados;
- Formatação financeira e percentual;
- Gráfico de composição de carteira;
- Git e GitHub para versionamento e documentação.

## Estrutura sugerida do repositório

```text
simulador-investimentos-fii/
├── README.md
├── Simulador_Investimentos_FII.xlsx
└── images/
    ├── simulador-fii.png
    └── tabela-perfis.png
```

## Aprendizados

Durante o desenvolvimento, foi possível compreender como transformar parâmetros informados pelo usuário em projeções financeiras automáticas. O projeto também permitiu praticar referências absolutas, funções financeiras, pesquisas em tabelas, validação de dados, formatação condicional e organização visual de uma planilha.

Outro aprendizado importante foi a documentação do projeto em Markdown e a organização dos arquivos em um repositório GitHub, facilitando a apresentação e o compartilhamento da solução.

## Observação

Este simulador possui finalidade exclusivamente educacional. Os resultados são estimativas baseadas nas taxas informadas e não representam garantia de rentabilidade nem recomendação de investimento.

## Autor

Desenvolvido por **Luis Henrique Gonçalves de Sá** para o desafio de Excel da DIO.
