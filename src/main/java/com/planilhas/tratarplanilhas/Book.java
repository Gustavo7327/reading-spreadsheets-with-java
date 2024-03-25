package com.planilhas.tratarplanilhas;

public class Book {

    private String titulo;
    private String autor;
    private String ano;
    private int exemplares;

    public Book(String titulo, String autor, String ano, int exemplares) {
        this.titulo = titulo;
        this.autor = autor;
        this.ano = ano;
        this.exemplares = exemplares;
    }

    public String getTitulo() {
        return titulo;
    }

    public String getAutor() {
        return autor;
    }

    public String getAno() {
        return ano;
    }

    public int getExemplares() {
        return exemplares;
    }

    public void setExemplares(int exemplares) {
        this.exemplares = exemplares;
    }
    
}
