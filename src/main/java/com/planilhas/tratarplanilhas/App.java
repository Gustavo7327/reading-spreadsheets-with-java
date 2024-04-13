package com.planilhas.tratarplanilhas;

import java.io.File;
import java.util.Map;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;

import com.aspose.cells.Cell;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.SaveFormat;

public class App 
{
    public static void main( String[] args )
    {
        Map<String, Book> biblioteca = Util.readSpreadsheet();

        Workbook wb = new Workbook();
        wb.getWorksheets().add("Biblioteca");

        Worksheet ws = wb.getWorksheets().get(0);
        ws.getCells().get(0, 0).setValue("TITULO");
        ws.getCells().get(0, 1).setValue("AUTOR");
        ws.getCells().get(0, 2).setValue("ANO");
        ws.getCells().get(0, 3).setValue("EXEMPLARES");

        int row = 1;

        for (String key : biblioteca.keySet()) {
            Book book = biblioteca.get(key);

            Cell titulo = ws.getCells().get(row, 0);
            titulo.setValue(book.getTitulo());

            Cell autor = ws.getCells().get(row, 1);
            autor.setValue(book.getAutor());

            Cell ano = ws.getCells().get(row, 2);
            ano.setValue(book.getAno());

            Cell exemplares = ws.getCells().get(row, 3);
            exemplares.setValue(book.getExemplares());

            row++;
        }

        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Save Workbook");
        FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel Files (*.xlsx)", "xlsx");
        fileChooser.setFileFilter(filter);

        int result = fileChooser.showSaveDialog(null);
        if(result == JFileChooser.APPROVE_OPTION){
            File file = fileChooser.getSelectedFile();
            try{
                wb.save(file.getAbsolutePath(), SaveFormat.XLSX);
                System.out.println("Workbook save in " + file.getAbsolutePath());
            } catch(Exception e){
                e.printStackTrace();
            }
        }            
    }
}
