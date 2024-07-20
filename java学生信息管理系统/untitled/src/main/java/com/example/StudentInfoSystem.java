package com.example;

import com.example.Student;
import com.example.StudentDAO;
import javafx.application.Application;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class StudentInfoSystem extends Application {

    private TableView<Student> table;
    private ObservableList<Student> data;

    @Override
    public void start(Stage primaryStage) {
        primaryStage.setTitle("学生信息管理系统");

        StudentDAO.createTable();

        TableColumn<Student, Integer> idCol = new TableColumn<>("ID");
        idCol.setCellValueFactory(new PropertyValueFactory<>("id"));

        TableColumn<Student, String> nameCol = new TableColumn<>("姓名");
        nameCol.setCellValueFactory(new PropertyValueFactory<>("name"));

        TableColumn<Student, Integer> ageCol = new TableColumn<>("年龄");
        ageCol.setCellValueFactory(new PropertyValueFactory<>("age"));

        TableColumn<Student, String> gradeCol = new TableColumn<>("年级");
        gradeCol.setCellValueFactory(new PropertyValueFactory<>("grade"));

        table = new TableView<>();
        table.getColumns().addAll(idCol, nameCol, ageCol, gradeCol);

        Button addButton = new Button("添加");
        addButton.setOnAction(e -> addStudent());

        Button updateButton = new Button("更新");
        updateButton.setOnAction(e -> updateStudent());

        Button deleteButton = new Button("删除");
        deleteButton.setOnAction(e -> deleteStudent());

        Button exportButton = new Button("导出到Excel");
        exportButton.setOnAction(e -> exportToExcel());

        VBox vbox = new VBox(10, table, addButton, updateButton, deleteButton, exportButton);
        vbox.setPadding(new Insets(10));

        Scene scene = new Scene(vbox, 600, 400);
        primaryStage.setScene(scene);
        primaryStage.show();

        loadTableData();
    }

    private void loadTableData() {
        List<Student> students = StudentDAO.getAllStudents();
        data = FXCollections.observableArrayList(students);
        table.setItems(data);
    }

    private void addStudent() {
        TextInputDialog nameDialog = new TextInputDialog();
        nameDialog.setTitle("添加学生");
        nameDialog.setHeaderText("请输入学生姓名:");
        String name = nameDialog.showAndWait().orElse("新学生");

        TextInputDialog ageDialog = new TextInputDialog("18");
        ageDialog.setTitle("添加学生");
        ageDialog.setHeaderText("请输入学生年龄:");
        int age = Integer.parseInt(ageDialog.showAndWait().orElse("18"));

        TextInputDialog gradeDialog = new TextInputDialog("一年级");
        gradeDialog.setTitle("添加学生");
        gradeDialog.setHeaderText("请输入学生年级:");
        String grade = gradeDialog.showAndWait().orElse("一年级");

        Student student = new Student(0, name, age, grade);
        StudentDAO.addStudent(student);
        loadTableData();
    }

    private void updateStudent() {
        Student selectedStudent = table.getSelectionModel().getSelectedItem();
        if (selectedStudent != null) {
            TextInputDialog nameDialog = new TextInputDialog(selectedStudent.getName());
            nameDialog.setTitle("更新学生");
            nameDialog.setHeaderText("更新学生姓名:");
            String name = nameDialog.showAndWait().orElse(selectedStudent.getName());

            TextInputDialog ageDialog = new TextInputDialog(String.valueOf(selectedStudent.getAge()));
            ageDialog.setTitle("更新学生");
            ageDialog.setHeaderText("更新学生年龄:");
            int age = Integer.parseInt(ageDialog.showAndWait().orElse(String.valueOf(selectedStudent.getAge())));

            TextInputDialog gradeDialog = new TextInputDialog(selectedStudent.getGrade());
            gradeDialog.setTitle("更新学生");
            gradeDialog.setHeaderText("更新学生年级:");
            String grade = gradeDialog.showAndWait().orElse(selectedStudent.getGrade());

            selectedStudent = new Student(selectedStudent.getId(), name, age, grade);
            StudentDAO.updateStudent(selectedStudent);
            loadTableData();
        }
    }

    private void deleteStudent() {
        Student selectedStudent = table.getSelectionModel().getSelectedItem();
        if (selectedStudent != null) {
            StudentDAO.deleteStudent(selectedStudent.getId());
            loadTableData();
        }
    }

    private void exportToExcel() {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("学生信息");

        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("ID");
        header.createCell(1).setCellValue("姓名");
        header.createCell(2).setCellValue("年龄");
        header.createCell(3).setCellValue("年级");

        int rowNum = 1;
        for (Student student : data) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(student.getId());
            row.createCell(1).setCellValue(student.getName());
            row.createCell(2).setCellValue(student.getAge());
            row.createCell(3).setCellValue(student.getGrade());
        }

        try (FileOutputStream fileOut = new FileOutputStream("学生信息.xlsx")) {
            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
        launch(args);
    }
}
