package com.yupi.yupicturebackend.excel;

import com.yupi.yupicturebackend.bean.Component;
import com.yupi.yupicturebackend.bean.ProcessResource;
import com.yupi.yupicturebackend.bean.Step;
import com.yupi.yupicturebackend.bean.Process;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ProcessExample {
        public static void main(String[] args) {
        // 创建测试数据
        List<Process> processes = createTestProcesses();
        
        // 导出到Excel
        String excelFilePath = "process_data.xlsx";
        exportProcessesToExcel(processes, excelFilePath);
    }
    
    /**
     * 创建测试数据
     */
    public static List<Process> createTestProcesses() {
        List<Process> processes = new ArrayList<>();
        
        // 创建第一个工序
        Process process1 = new Process();
        process1.setId("P001");
        
        Step step1 = createStep("S001", "C001", new String[]{"R001", "R002","2222","3231","R003", "R004","2222","3231"});
        Step step2 = createStep("S002", "C002", new String[]{"R003", "R004","2222","3231","R003", "R004","2222","3231","R003", "R004","2222","3231","3231"});
        Step step3 = createStep("S003", "C003", new String[]{"R003", "R004","R0022222"});
        
        process1.addStep(step1);
        process1.addStep(step2);
        process1.addStep(step3);
        
        // 创建第二个工序
        Process process2 = new Process();
        process2.setId("P002");
        
        Step step4 = createStep("S004", "C004", new String[]{"R005", "R006"});
        Step step5 = createStep("S005", "C005", new String[]{"R007", "R008", "R009"});
        
        process2.addStep(step4);
        process2.addStep(step5);
        
        processes.add(process1);
        processes.add(process2);
        
        return processes;
    }
    
    /**
     * 创建工步
     */
    public static Step createStep(String stepId, String componentId, String[] resourceIds) {
        Step step = new Step();
        step.setStepId(stepId);
        
        // 添加物料
        Component component = new Component();
        component.setId(componentId);
        step.addComponent(component);
        
        // 添加工艺参数
        for (String resourceId : resourceIds) {
            ProcessResource resource = new ProcessResource();
            resource.setResourceId(resourceId);
            step.addResource(resource);
        }
        
        return step;
    }
    

    
    /**
     * 将多个工序数据导出到Excel
     */
    public static void exportProcessesToExcel(List<Process> processes, String filePath) {
        try {
            // 读取指定位置excel（如果存在）
            Workbook workbook;
            File file = new File("/Users/qarar/code/code_back/yu-picture/yu-picture-backend/src/test/java/com/yupi/yupicturebackend/excel/组包装SOP模板.xlsx");
            FileInputStream inputStream = new FileInputStream(file);
            workbook = new XSSFWorkbook(inputStream);
            inputStream.close();

            // 为每个工序创建sheet页
            for (Process process : processes) {
                fillStepDataForProcess(workbook, process);
            }

            // 调整所有sheet的列宽
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                autoSizeColumns(workbook.getSheetAt(i));
            }

            // 保存文件
            saveWorkbook(workbook, filePath);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    
    /**
     * 将单个工序数据导出到Excel（保留原方法）
     */
    public static void exportProcessToExcel(Process process, String filePath) {
        try {
            Workbook workbook = new XSSFWorkbook();
            
            // 填充数据（支持分页）
            fillStepDataWithPagination(workbook, process);
            
            // 调整所有sheet的列宽
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                autoSizeColumns(workbook.getSheetAt(i));
            }
            
            // 保存文件
            saveWorkbook(workbook, filePath);
            
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    
    /**
     * 创建表头
     */
    private static void createHeader(Sheet sheet) {
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("工步");
        headerRow.createCell(1).setCellValue("物料");
        headerRow.createCell(2).setCellValue("工艺参数");
    }
    
    /**
     * 填充工步数据（支持分页）
     */
    private static void fillStepDataWithPagination(Workbook workbook, Process process) {
        int currentRow = 1;
        int pageRowLimit = 20; // 每页最多20行
        Sheet currentSheet = workbook.createSheet(process.getId());
        int sheetIndex = 0;
        
        // 创建第一个sheet的表头
        createHeader(currentSheet);
        
        for (Step step : process.getSteps()) {
            int maxRows = calculateMaxRows(step);
            
            // 检查是否需要创建新页面
            if (currentRow + maxRows > pageRowLimit) {
                // 创建新页面
                sheetIndex++;
                currentSheet = createNewSheet(workbook, process.getId(), sheetIndex);
                currentRow = 1; // 重置行号
            }
            
            // 合并单元格
            if (maxRows > 1) {
                mergeCells(currentSheet, currentRow, maxRows);
            }
            
            // 设置工步ID
            setStepId(currentSheet, currentRow, step);
            
            // 填充物料和工艺参数
            fillComponents(currentSheet, currentRow, step);
            fillResources(currentSheet, currentRow, step);
            
            currentRow += maxRows;
        }
    }
    
    /**
     * 为单个工序填充数据（支持分页）
     */
    private static void fillStepDataForProcess(Workbook workbook, Process process) {
        int currentRow = 1;
        int pageRowLimit = 20; // 每页最多20行
        Sheet currentSheet = workbook.createSheet(process.getId());
        int sheetIndex = 0;
        
        // 创建第一个sheet的表头
        createHeader(currentSheet);
        
        for (Step step : process.getSteps()) {
            int maxRows = calculateMaxRows(step);
            
            // 检查是否需要创建新页面
            if (currentRow + maxRows > pageRowLimit) {
                // 创建新页面
                sheetIndex++;
                currentSheet = createNewSheet(workbook, process.getId(), sheetIndex);
                currentRow = 1; // 重置行号
            }
            
            // 合并单元格
            if (maxRows > 1) {
                mergeCells(currentSheet, currentRow, maxRows);
            }
            
            // 设置工步ID
            setStepId(currentSheet, currentRow, step);
            
            // 填充物料和工艺参数
            fillComponents(currentSheet, currentRow, step);
            fillResources(currentSheet, currentRow, step);
            
            currentRow += maxRows;
        }
    }
    
    /**
     * 创建新页面
     */
    private static Sheet createNewSheet(Workbook workbook, String processId, int sheetIndex) {
        String sheetName = processId + "_" + sheetIndex;
        Sheet newSheet = workbook.createSheet(sheetName);
        
        // 创建表头
        createHeader(newSheet);
        
        return newSheet;
    }
    
    /**
     * 计算工步需要的最大行数
     */
    private static int calculateMaxRows(Step step) {
        int maxRows = Math.max(step.getComponents().size(), step.getResources().size());
        return Math.max(maxRows, 1);
    }
    
    /**
     * 合并单元格
     */
    private static void mergeCells(Sheet sheet, int currentRow, int maxRows) {
        CellRangeAddress mergedRegion = new CellRangeAddress(currentRow, currentRow + maxRows - 1, 0, 2);
        sheet.addMergedRegion(mergedRegion);
    }
    
    /**
     * 设置工步ID
     */
    private static void setStepId(Sheet sheet, int currentRow, Step step) {
        Row stepRow = sheet.createRow(currentRow);
        Cell stepCell = stepRow.createCell(0);
        stepCell.setCellValue(step.getStepId());
    }
    
    /**
     * 填充物料数据
     */
    private static void fillComponents(Sheet sheet, int currentRow, Step step) {
        for (int i = 0; i < step.getComponents().size(); i++) {
            Row row = sheet.getRow(currentRow + i);
            if (row == null) {
                row = sheet.createRow(currentRow + i);
            }
            row.createCell(2).setCellValue(step.getComponents().get(i).getId());
        }
    }
    
    /**
     * 填充工艺参数数据
     */
    private static void fillResources(Sheet sheet, int currentRow, Step step) {
        for (int i = 0; i < step.getResources().size(); i++) {
            Row row = sheet.getRow(currentRow + i);
            if (row == null) {
                row = sheet.createRow(currentRow + i);
            }
            row.createCell(3).setCellValue(step.getResources().get(i).getResourceId());
        }
    }
    
    /**
     * 自动调整列宽
     */
    private static void autoSizeColumns(Sheet sheet) {
        for (int i = 0; i < 4; i++) {
            sheet.autoSizeColumn(i);
        }
    }
    
    /**
     * 保存工作簿
     */
    private static void saveWorkbook(Workbook workbook, String filePath) throws IOException {
        try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
            workbook.write(outputStream);
        }
        workbook.close();
    }
}
