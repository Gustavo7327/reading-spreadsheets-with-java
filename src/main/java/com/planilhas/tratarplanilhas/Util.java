package com.planilhas.tratarplanilhas;

import java.util.HashMap;
import java.util.Map;

import com.aspose.cells.Cell;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class Util {

    private static String createKey(String autor, String titulo, String ano){
        return autor + "-" + titulo + "-" + ano;
    }

    public static Map<String,Book> readSpreadsheet(){

        Map<String, Book> biblioteca = new HashMap<>();

        try {

            //carrega a planilha
            Workbook wb = new Workbook("/home/guilherme/Documentos/ProjetoBibliotecaUtils/library.xlsx");
            Worksheet ws = wb.getWorksheets().get(0);
            //System.out.printf("Planilha obtida: %s \n", ws.getName());

            //especificando cada coluna
            Cell copyNumberCell = ws.getCells().get(0, 3);
            copyNumberCell.setValue("EXEMPLARES");
            Cell autorCell = ws.getCells().get(0, 0);
            autorCell.setValue("AUTOR");
            Cell tituloCell = ws.getCells().get(0, 1);
            tituloCell.setValue("TITULO");
            Cell anoCell = ws.getCells().get(0, 2);
            anoCell.setValue("ANO");
            //System.out.println(copyNumberCell.getValue());

            for(int row = 1; row <= ws.getCells().getMaxDataRow(); row++){

                String autor = ws.getCells().get(row, 0).getStringValue();
                String titulo = ws.getCells().get(row, 1).getStringValue();
                String ano = !ws.getCells().get(row, 2).getStringValue().isEmpty() ? ano = String.valueOf(ws.getCells().get(row, 2).getStringValue()) : "NI";
                
                String key = createKey(autor, titulo, ano);
                Book livroExistente = biblioteca.get(key);

                if(livroExistente != null){
                    livroExistente.setExemplares(livroExistente.getExemplares()+1);
                }
                else{
                    Book novoLivro = new Book(titulo, autor, ano, 1);
                    biblioteca.put(key, novoLivro);
                }
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
        return biblioteca;
    }
}
