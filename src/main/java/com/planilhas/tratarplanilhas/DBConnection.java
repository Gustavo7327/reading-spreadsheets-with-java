package com.planilhas.tratarplanilhas;
import java.sql.*;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class DBConnection {
    public static void main(String[] args) {
        Connection conexao = null;
        String driverName = "com.mysql.cj.jdbc.Driver";
        
        try {
            Class.forName(driverName);
        } catch (ClassNotFoundException e) {
            e.printStackTrace();
        }

        String url = "jdbc:mysql://localhost:3306/library";
        String user = "gustavo7327";
        String password = "gustavin7327";

        try{
            conexao = DriverManager.getConnection(url, user, password);
            System.out.println("Conectado ao banco de dados!");

            Workbook wb = new Workbook("/home/guilherme/Documentos/ProjetoBibliotecaUtils/BibliotecaPlanilha.xlsx");
            Worksheet ws = wb.getWorksheets().get("DadosRequisitados");

            loop(ws, conexao);

            conexao.close();
            System.out.println("Conexão fechada!");
          } catch (SQLException e) {
                e.printStackTrace();
            }
            catch (Exception e){
                e.printStackTrace();
            }
    }

    public static void insert(Connection conn, String titulo, String autores, String generos, int ano_publicacao, String url_imagem, String url_amostra, String sinopse, double avaliacao, int numero_paginas, int numero_exemplares){

        String query = "insert into books (titulo, autores, generos, ano_publicacao, url_imagem, url_amostra, sinopse, avaliacao, numero_paginas, numero_exemplares)" + " values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";

        try {

            PreparedStatement preparedStatement = conn.prepareStatement(query);
            preparedStatement.setString(1, titulo);
            preparedStatement.setString(2, autores);
            preparedStatement.setString(3, generos);
            preparedStatement.setInt(4, ano_publicacao);
            preparedStatement.setString(5, url_imagem);
            preparedStatement.setString(6, url_amostra);
            preparedStatement.setString(7, sinopse);
            preparedStatement.setDouble(8, avaliacao);
            preparedStatement.setInt(9, numero_paginas);
            preparedStatement.setInt(10, numero_exemplares);

            preparedStatement.execute();

        } catch (SQLException e) {
            e.printStackTrace();
        }
    }

    public static void loop(Worksheet ws, Connection conn){
        for(int i = 1; i < ws.getCells().getMaxDataRow(); i++){

            String titulo = ws.getCells().get(i, 0).getStringValue();
            String autores = ws.getCells().get(i, 1).getStringValue();
            int ano_publicacao = Integer.parseInt(ws.getCells().get(i, 2).getStringValue());
            String url_imagem = ws.getCells().get(i, 3).getStringValue();
            String url_amostra = ws.getCells().get(i, 4).getStringValue();
            String sinopse = ws.getCells().get(i, 5).getStringValue();
            String generos = ws.getCells().get(i, 6).getStringValue();
            double avaliacao = Double.parseDouble(ws.getCells().get(i, 7).getStringValue());
            int numero_paginas = Integer.parseInt(ws.getCells().get(i, 8).getStringValue());
            int numero_exemplares = Integer.parseInt(ws.getCells().get(i, 9).getStringValue());

            insert(conn, titulo, autores, generos, ano_publicacao, url_imagem, url_amostra, sinopse, avaliacao, numero_paginas, numero_exemplares);
        }
    }

}
