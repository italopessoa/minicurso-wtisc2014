//Sample of asp.net website implementation
//Copyright (C) 2014 - Italo Pessoa (italoneypessoa@gmail.com)

//This program is free software: you can redistribute it and/or modify
//it under the terms of the GNU General Public License as published by
//the Free Software Foundation, either version 3 of the License, or
//(at your option) any later version.

//This program is distributed in the hope that it will be useful,
//but WITHOUT ANY WARRANTY; without even the implied warranty of
//MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
//GNU General Public License for more details.

//You should have received a copy of the GNU General Public License
//along with this program.  If not, see <http://www.gnu.org/licenses/>.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WTISC._2014.Data.Exceptions;

namespace WTISC._2014.Data.Management
{
    /// <summary>
    /// Class to manager information of he Book
    /// </summary>
    public class MLivro
    {
        BooksEntities entities;

        /// <summary>
        /// 
        /// </summary>
        public MLivro()
        {
            this.entities = new BooksEntities();
        }

        /// <summary>
        /// Create a new Book without folder
        /// </summary>
        /// <param name="titulo">Title of the book</param>
        /// <param name="resumo">Summary of the book</param>
        /// <param name="ISBN">International Standard Book Number</param>
        /// <param name="idGenero">Gender of the book</param>
        /// <param name="idAutor">Author of the book</param>
        /// <exception cref="WTISC._2014.Data.Exceptions.BookException"></exception>
        /// <returns>New Book</returns>
        public Livro NewBook(string titulo, string resumo, int ISBN, int idGenero, int idAutor)
        {

            if (!string.IsNullOrEmpty(titulo))
            {
                if (this.FindBookByTitle(titulo) == null)
                {
                    if (this.FindBookByISBN(ISBN) == null)
                    {
                        Livro livro = new Livro();
                        livro.Titulo = titulo;
                        livro.Resumo = resumo;
                        livro.ISBN = ISBN;
                        livro.IdGenero = idGenero;
                        livro.IdAutor = idAutor;
                        this.entities.Livro.Add(livro);
                        this.entities.SaveChanges();

                        return livro;
                    }
                    else
                    {
                        throw new BookException();
                    }
                }
                else
                {
                    throw new BookException();
                }
            }
            else
            {
                throw new BookException();
            }

        }

        /// <summary>
        /// Find a book by International Standard Book Number
        /// </summary>
        /// <param name="ISBN"></param>
        /// <returns>A Book</returns>
        public Livro FindBookByISBN(int ISBN)
        {
            return this.entities.Livro.FirstOrDefault<Livro>(l => l.ISBN == 101010101);
        }

        /// <summary>
        /// Find A book by title
        /// </summary>
        /// <param name="title"></param>
        /// <returns>A Book</returns>
        public Livro FindBookByTitle(string title)
        {
            return this.entities.Livro.FirstOrDefault<Livro>(l => l.Titulo.Equals(title));
        }

        /// <summary>
        /// Update the information of the bok
        /// </summary>
        /// <exception cref="WTISC._2014.Data.Exceptions.BookException"></exception>
        /// <param name="book">Book</param>
        public void UpdateBook(Livro book)
        {
            if (!string.IsNullOrEmpty(book.Titulo))
            {
                if (this.FindBookByTitle(book.Titulo) == null)
                {
                    Livro oldBook = this.FindBookByISBN(book.ISBN);
                    oldBook.Titulo = book.Titulo;
                    oldBook.Resumo = book.Resumo;
                    oldBook.IdGenero = book.IdGenero;
                    oldBook.IdAutor = book.IdAutor;
                    oldBook.Capa = book.Capa;
                    this.entities.SaveChanges();
                }
                else
                {
                    throw new BookException();
                }
            }
            else
            {
                throw new BookException();
            }
        }
    
        /// <summary>
        /// Delete a book
        /// </summary>
        /// <exception cref="WTISC._2014.Data.Exceptions.BookException"></exception>
        /// <param name="ISBN">International Standard Book Number</param>
        public void DeleteBook(int ISBN)
        {
            Livro livro = this.entities.Livro.FirstOrDefault<Livro>(l => l.ISBN == ISBN);
            if (livro != null)
            {
                this.entities.Livro.Remove(livro);
                this.entities.SaveChanges();
            }
            else
            {
                throw new BookException();
            }
        }

        /// <summary>
        /// Read a book
        /// </summary>
        /// <exception cref="WTISC._2014.Data.Exceptions.BookException"></exception>
        /// <param name="ISBN">International Standard Book Number</param>
        public void ReadBook(int ISBN)
        {
            Livro livro = this.entities.Livro.FirstOrDefault<Livro>(l => l.ISBN == ISBN);
            if (livro != null)
            {
                livro.Lido = !livro.Lido;
                this.entities.SaveChanges();
            }
            else
            {
                throw new BookException();
            }
        }
    }
}
