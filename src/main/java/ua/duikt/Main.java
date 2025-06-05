package ua.duikt;

import javafx.application.Application;
import javafx.beans.property.SimpleStringProperty;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.SelectionMode;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.usermodel.*;
import ua.duikt.entity.Bachelor;
import ua.duikt.repository.BachelorRepository;
import ua.duikt.util.RunSnapshot;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

@Slf4j
public class Main extends Application {

    private final BachelorRepository bachelorRepository = new BachelorRepository();

    @Override
    public void start(Stage stage) {
        TableView<Bachelor> table = new TableView<>();
        table.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);

        TableColumn<Bachelor, String> nameCol = new TableColumn<>("Bachelor Full Name");
        nameCol.setCellValueFactory(data -> new SimpleStringProperty(data.getValue().getFullName()));

        TableColumn<Bachelor, String> masterNameCol = new TableColumn<>("Master Full Name");
        masterNameCol.setCellValueFactory(data -> new SimpleStringProperty(data.getValue().getMasterFullName()));

        TableColumn<Bachelor, String> topicCol = new TableColumn<>("Topic");
        topicCol.setCellValueFactory(data -> new SimpleStringProperty(data.getValue().getTopic()));

        table.getColumns().addAll(nameCol, masterNameCol, topicCol);
        table.getItems().addAll(bachelorRepository.getAll());

        Button exportBtn = new Button("Export to DOCX");
        exportBtn.setOnAction(e -> exportToDocx(table.getSelectionModel().getSelectedItems()));

        VBox root = new VBox(10, table, exportBtn);
        Scene scene = new Scene(root, 600, 400);
        stage.setTitle("Person Exporter");
        stage.setScene(scene);
        stage.show();
    }

    private void exportToDocx(List<Bachelor> people) {
        for (Bachelor b : people) {
            try (XWPFDocument doc = new XWPFDocument(Objects.requireNonNull(getClass().getResourceAsStream("/Форма_шаблон.docx")))) {

                for (XWPFParagraph paragraph : doc.getParagraphs()) {
                    replacePlaceholdersInParagraph(paragraph, b);
                }

                for (XWPFTable table : doc.getTables()) {
                    for (XWPFTableRow row : table.getRows()) {
                        for (XWPFTableCell cell : row.getTableCells()) {
                            for (XWPFParagraph paragraph : cell.getParagraphs()) {
                                replacePlaceholdersInParagraph(paragraph, b);
                            }
                        }
                    }
                }

                String fileName = b.getFullName().replace(" ", "_") + ".docx";
                try (FileOutputStream out = new FileOutputStream(fileName)) {
                    doc.write(out);
                }

                log.info("Saved file: {}", fileName);

            } catch (Exception ex) {
                log.error("Error exporting for {}: {}", b.getFullName(), ex.getMessage(), ex);
            }
        }
    }

    private void replacePlaceholdersInParagraph(XWPFParagraph paragraph, Bachelor b) {
        List<XWPFRun> runs = paragraph.getRuns();
        if (runs == null || runs.isEmpty()) return;

        StringBuilder fullTextBuilder = new StringBuilder();
        List<RunSnapshot> styleSnapshots = new ArrayList<>();

        for (XWPFRun run : runs) {
            String text = run.getText(0);
            if (text != null) {
                fullTextBuilder.append(text);
                styleSnapshots.add(new RunSnapshot(run));
            }
        }

        String fullText = fullTextBuilder.toString();

        String replaced = fullText
                .replace("${FULL_NAME}", b.getFullName())
                .replace("${MASTER_NAME}", b.getMasterFullName())
                .replace("${TOPIC}", b.getTopic());

        for (int i = runs.size() - 1; i >= 0; i--) {
            paragraph.removeRun(i);
        }

        XWPFRun newRun = paragraph.createRun();
        if (!styleSnapshots.isEmpty()) {
            styleSnapshots.get(0).applyTo(newRun);
        } else {
            newRun.setFontFamily("Times New Roman");
            newRun.setFontSize(14);
        }
        newRun.setText(replaced);
    }

    public static void main(String[] args) {
        launch(args);
    }
}