package com.salonso;

import javafx.application.Application;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.chart.BarChart;
import javafx.scene.chart.CategoryAxis;
import javafx.scene.chart.NumberAxis;
import javafx.scene.chart.XYChart;
import javafx.scene.control.*;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;

import java.io.File;
import java.io.FileInputStream;
import java.util.*;

public class ExcelDashboard extends Application {

    private BarChart<String, Number> barChart;
    private Label statusLabel;
    private Label totalLabel;
    private Label statsLabel;
    private ComboBox<String> columnSelector;
    private ComboBox<String> relationColumnSelector;
    private TextField filterField;
    private List<Map<String, String>> dados;
    private List<String> headers;
    private CheckBox showAllDataCheckBox;
    private CheckBox showPercentageCheckBox;
    private Button relationButton;

    @Override
    public void start(Stage stage) {
        VBox root = new VBox(10);
        root.setPadding(new Insets(10));

        Button selectFileButton = new Button("Selecionar Planilha Excel");
        statusLabel = new Label("Nenhum arquivo selecionado");
        totalLabel = new Label("Total: 0");
        statsLabel = new Label("");

        GridPane selectors = new GridPane();
        selectors.setHgap(10);
        selectors.setVgap(10);

        columnSelector = new ComboBox<>();
        relationColumnSelector = new ComboBox<>();
        filterField = new TextField();
        filterField.setPromptText("Filtrar por valor...");

        showAllDataCheckBox = new CheckBox("Mostrar todos os dados");
        showPercentageCheckBox = new CheckBox("Mostrar porcentagens");
        relationButton = new Button("Relacionar Colunas");

        selectors.add(new Label("Coluna Principal:"), 0, 0);
        selectors.add(columnSelector, 1, 0);
        selectors.add(new Label("Coluna Relacionada:"), 2, 0);
        selectors.add(relationColumnSelector, 3, 0);
        selectors.add(new Label("Filtro:"), 0, 1);
        selectors.add(filterField, 1, 1);
        selectors.add(showAllDataCheckBox, 2, 1);
        selectors.add(showPercentageCheckBox, 3, 1);
        selectors.add(relationButton, 4, 1);

        CategoryAxis xAxis = new CategoryAxis();
        NumberAxis yAxis = new NumberAxis();
        barChart = new BarChart<>(xAxis, yAxis);
        barChart.setTitle("Dashboard de Dados");

        selectFileButton.setOnAction(e -> {
            FileChooser fileChooser = new FileChooser();
            fileChooser.getExtensionFilters().add(
                    new FileChooser.ExtensionFilter("Excel Files", "*.xlsx")
            );
            File file = fileChooser.showOpenDialog(stage);
            if (file != null) {
                loadExcelData(file);
            }
        });

        columnSelector.setOnAction(e -> updateChart());
        filterField.setOnAction(e -> updateChart());
        showAllDataCheckBox.setOnAction(e -> updateChart());
        showPercentageCheckBox.setOnAction(e -> updateChart());
        relationButton.setOnAction(e -> updateRelationChart());

        HBox statsBox = new HBox(20);
        statsBox.getChildren().addAll(totalLabel, statsLabel);

        root.getChildren().addAll(selectFileButton, statusLabel, selectors, statsBox, barChart);

        Scene scene = new Scene(root, 1000, 800);
        stage.setTitle("Dashboard Excel");
        stage.setScene(scene);
        stage.show();
    }

    private void loadExcelData(File file) {
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            dados = new ArrayList<>();
            headers = new ArrayList<>();

            Row headerRow = sheet.getRow(0);
            if (headerRow != null) {
                for (org.apache.poi.ss.usermodel.Cell cell : headerRow) {
                    if (cell != null) {
                        String header = cell.getStringCellValue();
                        headers.add(header);
                    }
                }
            }

            columnSelector.getItems().clear();
            relationColumnSelector.getItems().clear();
            columnSelector.getItems().addAll(headers);
            relationColumnSelector.getItems().addAll(headers);

            if (!headers.isEmpty()) {
                columnSelector.setValue(headers.get(0));
                if (headers.size() > 1) {
                    relationColumnSelector.setValue(headers.get(1));
                }
            }

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    Map<String, String> rowData = new HashMap<>();
                    for (int j = 0; j < headers.size(); j++) {
                        org.apache.poi.ss.usermodel.Cell cell = row.getCell(j);
                        String value = "";
                        if (cell != null) {
                            switch (cell.getCellType()) {
                                case NUMERIC:
                                    value = String.format("%.2f", cell.getNumericCellValue());
                                    break;
                                case STRING:
                                    value = cell.getStringCellValue();
                                    break;
                                default:
                                    value = "";
                            }
                        }
                        rowData.put(headers.get(j), value);
                    }
                    dados.add(rowData);
                }
            }

            statusLabel.setText("Arquivo carregado: " + file.getName());
            updateChart();

        } catch (Exception e) {
            statusLabel.setText("Erro ao carregar arquivo: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private void updateChart() {
        if (dados == null || dados.isEmpty() || columnSelector.getValue() == null)
            return;

        barChart.getData().clear();
        XYChart.Series<String, Number> series = new XYChart.Series<>();
        series.setName("Dados");

        String filterText = filterField.getText().toLowerCase();
        String selectedColumn = columnSelector.getValue();

        Map<String, Integer> countMap = new HashMap<>();
        int totalCount = 0;

        for (Map<String, String> row : dados) {
            String value = row.get(selectedColumn);
            if (value == null || value.trim().isEmpty()) continue;

            if (!filterText.isEmpty() && !value.toLowerCase().contains(filterText)) {
                continue;
            }

            countMap.merge(value, 1, Integer::sum);
            totalCount++;
        }

        totalLabel.setText(String.format("Total: %d", totalCount));

        List<Map.Entry<String, Integer>> sortedEntries = new ArrayList<>(countMap.entrySet());
        sortedEntries.sort(Map.Entry.<String, Integer>comparingByValue().reversed());

        int limit = showAllDataCheckBox.isSelected() ? sortedEntries.size() : Math.min(10, sortedEntries.size());

        String[] colors = {
                "#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd",
                "#8c564b", "#e377c2", "#7f7f7f", "#bcbd22", "#17becf"
        };

        for (int i = 0; i < limit; i++) {
            Map.Entry<String, Integer> entry = sortedEntries.get(i);
            double percentage = (entry.getValue() * 100.0) / totalCount;
            String label = showPercentageCheckBox.isSelected() ?
                    String.format("%s (%.1f%%)", entry.getKey(), percentage) :
                    entry.getKey();
            XYChart.Data<String, Number> data = new XYChart.Data<>(label, entry.getValue());
            series.getData().add(data);

            String color = colors[i % colors.length];
            data.getNode().setStyle(String.format("-fx-bar-fill: %s;", color));
        }

        barChart.getData().add(series);
    }

    private void updateRelationChart() {
        String col1 = columnSelector.getValue();
        String col2 = relationColumnSelector.getValue();

        if (col1 == null || col2 == null || col1.equals(col2)) {
            statusLabel.setText("Selecione duas colunas diferentes para relacionar");
            return;
        }

        Set<String> uniqueValues1 = new TreeSet<>();
        Set<String> uniqueValues2 = new TreeSet<>();
        Map<String, Map<String, Integer>> relationMap = new HashMap<>();

        for (Map<String, String> row : dados) {
            String val1 = row.get(col1);
            String val2 = row.get(col2);

            if (val1 != null && !val1.trim().isEmpty()) uniqueValues1.add(val1.trim());
            if (val2 != null && !val2.trim().isEmpty()) uniqueValues2.add(val2.trim());
        }

        for (String val1 : uniqueValues1) {
            Map<String, Integer> innerMap = new HashMap<>();
            for (String val2 : uniqueValues2) {
                innerMap.put(val2, 0);
            }
            relationMap.put(val1, innerMap);
        }

        for (Map<String, String> row : dados) {
            String val1 = row.get(col1);
            String val2 = row.get(col2);

            if (val1 != null && !val1.trim().isEmpty() &&
                    val2 != null && !val2.trim().isEmpty()) {
                val1 = val1.trim();
                val2 = val2.trim();
                relationMap.get(val1).merge(val2, 1, Integer::sum);
            }
        }

        barChart.getData().clear();
        String description = String.format(
                "Relação entre '%s' (cores) e '%s' (categorias)\n" +
                        "Cada cor representa um valor de '%s' e mostra sua quantidade em cada '%s'",
                col1, col2, col1, col2
        );
        barChart.setTitle(description);

        String[] colors = {
                "#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd",
                "#8c564b", "#e377c2", "#7f7f7f", "#bcbd22", "#17becf"
        };

        List<String> limitedValues1 = new ArrayList<>(uniqueValues1);
        if (!showAllDataCheckBox.isSelected() && limitedValues1.size() > 10) {
            limitedValues1 = limitedValues1.subList(0, 10);
        }

        Map<String, String> colorMap = new HashMap<>();
        int colorIndex = 0;
        for (String val1 : limitedValues1) {
            colorMap.put(val1, colors[colorIndex % colors.length]);
            colorIndex++;
        }

        for (String val1 : limitedValues1) {
            XYChart.Series<String, Number> series = new XYChart.Series<>();

            Map<String, Integer> innerMap = relationMap.get(val1);
            int total = innerMap.values().stream().mapToInt(Integer::intValue).sum();
            String seriesName = showPercentageCheckBox.isSelected() ?
                    String.format("%s (Total: %d)", val1, total) : val1;
            series.setName(seriesName);

            for (String val2 : uniqueValues2) {
                int count = innerMap.get(val2);
                if (count > 0) {
                    String label = val2;
                    if (showPercentageCheckBox.isSelected()) {
                        double percentage = (count * 100.0) / total;
                        label = String.format("%s (%.1f%%)", val2, percentage);
                    }
                    series.getData().add(new XYChart.Data<>(label, count));
                }
            }

            if (!series.getData().isEmpty()) {
                barChart.getData().add(series);
                String color = colorMap.get(val1);
                series.getData().forEach(data ->
                        data.getNode().setStyle(String.format("-fx-bar-fill: %s;", color))
                );
            }
        }

        int totalRelations = relationMap.values().stream()
                .flatMap(m -> m.values().stream())
                .mapToInt(Integer::intValue)
                .sum();

        statsLabel.setText(String.format(
                "Total de relações: %d\n" +
                        "Valores únicos em '%s': %d\n" +
                        "Valores únicos em '%s': %d",
                totalRelations, col1, uniqueValues1.size(), col2, uniqueValues2.size()
        ));

        totalLabel.setText(String.format("Total geral: %d", totalRelations));
    }

    public static void main(String[] args) {
        launch(args);
    }
}