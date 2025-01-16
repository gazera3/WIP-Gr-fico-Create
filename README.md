# Dashboard de Análise de Planilhas Excel

Aplicativo Java para visualização e análise de dados de planilhas Excel com interface gráfica JavaFX.

## Recursos

- Leitura de planilhas Excel (.xlsx)
- Visualização em gráficos de barras
- Filtragem de dados
- Relação entre colunas
- Estatísticas básicas
- Porcentagens e totais

## Requisitos

- Java 21
- Maven
- JavaFX 21
- Apache POI 5.2.3

## Instalação

```bash
git clone https://github.com/gazera3/WIP-Grafico-Create
cd planilhas
mvn clean install
```

## Execução

```bash
mvn javafx:run
```

## Uso

1. Clique em "Selecionar Planilha Excel"
2. Escolha as colunas para análise
3. Use os filtros e checkboxes para personalizar a visualização
4. Para relacionar colunas, selecione duas colunas diferentes e clique em "Relacionar Colunas"

## Estrutura do Projeto

```
src/
├── main/
│   ├── java/
│   │   └── com/
│   │       └── salonso/
│   │           └── ExcelDashboard.java
│   └── resources/
└── pom.xml
```

## Autora

Sophie Alonso

## Licença

Este projeto está licenciado sob a MIT License
