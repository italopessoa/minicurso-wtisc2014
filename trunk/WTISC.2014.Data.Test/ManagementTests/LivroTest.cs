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
using Microsoft.VisualStudio.TestTools.UnitTesting;
using WTISC._2014.Data.Management;
using WTISC._2014.Data.Exceptions;
using System.Collections.Generic;

namespace WTISC._2014.Data.Test.ManagementTests
{
    [TestClass]
    public class LivroTest
    {
        #region Private Members

        private MLivro managementTest;
        private const string TITLE = "title test";
        private const string DESCRIPTION = "book description test";
        private const int ISBN = 123456789;
        private int idGender;
        private int idAuthor;

        #endregion

        #region Livro Tests

        [TestInitialize]
        [Description("Initialize test")]
        private void Initialize()
        {
            this.managementTest = new MLivro();
            MAutor mAuthor = new MAutor();
            MGenero mGender = new MGenero();
            idAuthor = mAuthor.NewAuthor("author test").Id;
            idGender = mGender.NewGender("gender test", "description").Id;
        }

        [TestCleanup]
        [Description("Delete database values")]
        private void RestartDB()
        {
            BooksEntities resetEntities = new BooksEntities();
            resetEntities.Database.ExecuteSqlCommand("DELETE [Livro]");
            resetEntities.Database.ExecuteSqlCommand("DELETE [Autor]");
            resetEntities.Database.ExecuteSqlCommand("DELETE [Genero]");
        }

        [TestMethod]
        [Description("Create a book")]
        public void NewBookTest()
        {
            #region Case 1: Create a new Book

            Livro expected = new Livro();
            expected.Titulo = TITLE;
            expected.Resumo = DESCRIPTION;
            expected.IdAutor = idAuthor;
            expected.IdGenero = idGender;
            expected.ISBN = ISBN;

            Livro actual = this.managementTest.NewBook(TITLE, DESCRIPTION, ISBN, idGender, idAuthor);

            Assert.AreEqual(expected, actual);

            #endregion
        }

        [TestMethod]
        [ExpectedException(typeof(BookException), "A book with the same ISBN code already exists!")]
        [Description("Create a book with a already existing  ISBN")]
        public void NewBookException_ISBNTest()
        {
            #region Case 1: Create a book with already existing ISBN

            this.managementTest.NewBook(TITLE + "qwe", DESCRIPTION + "qwe", ISBN, idGender, idAuthor);
            this.managementTest.NewBook(TITLE, DESCRIPTION, ISBN, idGender, idAuthor);

            #endregion
        }

        [TestMethod]
        [ExpectedException(typeof(BookException), "The title can't be empty!")]
        [Description("Create a book without title")]
        public void NewBookException_EmptyTitleTest()
        {
            #region Case 1: Create a book without title

            this.managementTest.NewBook(string.Empty, DESCRIPTION + "qwe", ISBN, idGender, idAuthor);

            #endregion
        }

        [TestMethod]
        [Description("Find a book by ISBN")]
        public void FindBookByISBNTest()
        {
            #region Case 1: Find a book by the ISBN code

            Livro expected = new Livro();
            expected.Titulo = TITLE;
            expected.Resumo = DESCRIPTION;
            expected.IdAutor = idAuthor;
            expected.IdGenero = idGender;
            expected.ISBN = ISBN;

            this.managementTest.NewBook(TITLE, DESCRIPTION, ISBN, idGender, idAuthor);
            Livro actual = this.managementTest.FindBookByISBN(ISBN);
            Assert.AreEqual(expected, actual);

            #endregion

            #region Case 2: No result

            expected = null;

            actual = this.managementTest.FindBookByISBN(0987654321);
            Assert.AreEqual(expected, actual);

            #endregion
        }

        [TestMethod]
        [Description("Find a book by title")]
        public void FindBookByTitleTest()
        {
            #region Case 1: Find a book by the title

            Livro expected = new Livro();
            expected.Titulo = TITLE;
            expected.Resumo = DESCRIPTION;
            expected.IdAutor = idAuthor;
            expected.IdGenero = idGender;
            expected.ISBN = ISBN;

            this.managementTest.NewBook(TITLE, DESCRIPTION, ISBN, idGender, idAuthor);
            Livro actual = this.managementTest.FindBookByTitle(TITLE);
            Assert.AreEqual(expected, actual);

            #endregion

            #region Case 2: No result

            expected = null;

            actual = this.managementTest.FindBookByTitle("invalide title");
            Assert.AreEqual(expected, actual);

            #endregion
        }

        [TestMethod]
        [Description("Alter book values")]
        public void UpdateBookTeste()
        {
            #region Case 1: Alter a book

            Livro expected = new Livro();
            expected.Titulo = TITLE;
            expected.Resumo = DESCRIPTION;
            expected.IdAutor = idAuthor;
            expected.IdGenero = idGender;
            expected.ISBN = ISBN;

            expected = this.managementTest.NewBook(TITLE, DESCRIPTION, ISBN, idGender, idAuthor);

            expected.Titulo = "NEW TITLE";

            Livro actual = this.managementTest.FindBookByISBN(expected.ISBN);
            Assert.AreEqual(expected, actual);

            #endregion
        }
    
        [TestMethod]
        [ExpectedException(typeof(BookException),"The title can't be empty!")]
        [Description("Alter book settin an empty title")]
        public void UpdateBookException_EmptyTitleTest()
        {
            #region Case 1: Alter a book without title

            Livro book = new Livro();
            book.Titulo = TITLE;
            book.Resumo = DESCRIPTION;
            book.IdAutor = idAuthor;
            book.IdGenero = idGender;
            book.ISBN = ISBN;

            book = this.managementTest.NewBook(TITLE, DESCRIPTION, ISBN, idGender, idAuthor);

            book.Titulo = String.Empty; ;

            this.managementTest.UpdateBook(book);

            #endregion
        }

        [TestMethod]
        [Description("Delete the book")]
        public void DeleteBookTest()
        {
            #region Case 1: Delete a book

            Livro book = new Livro();
            book.Titulo = TITLE;
            book.Resumo = DESCRIPTION;
            book.IdAutor = idAuthor;
            book.IdGenero = idGender;
            book.ISBN = ISBN;

            book = this.managementTest.NewBook(TITLE, DESCRIPTION, ISBN, idGender, idAuthor);

            this.managementTest.DeleteBook(book.ISBN);

            Livro actual = this.managementTest.FindBookByISBN(book.ISBN);
            Assert.AreEqual(null, actual);

            #endregion
        }

        [TestMethod]
        [ExpectedException(typeof(BookException), "Book not found!")]
        [Description("Find an inexistent book")]
        public void DeleteBookException_NoResultTest()
        {
            this.managementTest.DeleteBook(958575);
        }

        [TestMethod]
        [Description("Change the value 'Lido'")]
        public void ReadBookTest()
        {
            #region Case 1: Read book

            bool expected = true;
            Livro book = new Livro();
            book.Titulo = TITLE;
            book.Resumo = DESCRIPTION;
            book.IdAutor = idAuthor;
            book.IdGenero = idGender;
            book.ISBN = ISBN;

            book = this.managementTest.NewBook(TITLE, DESCRIPTION, ISBN, idGender, idAuthor);

            this.managementTest.ReadBook(book.ISBN);

            book = this.managementTest.FindBookByISBN(book.ISBN);
            
            Assert.AreEqual(expected, book.Lido);

            #endregion

            #region Case 2: Unread book

            expected = false;

            this.managementTest.ReadBook(book.ISBN);

            book = this.managementTest.FindBookByISBN(book.ISBN);

            Assert.AreEqual(expected, book.Lido);

            #endregion
        }

        [TestMethod]
        [ExpectedException(typeof(BookException), "Book not found!")]
        [Description("Read an inexistent book")]
        public void ReadBookExceptionTest()
        {
            #region Case 1: Book not found

            Livro book = new Livro();
            book.Titulo = TITLE;
            book.Resumo = DESCRIPTION;
            book.IdAutor = idAuthor;
            book.IdGenero = idGender;
            book.ISBN = ISBN;

            book = this.managementTest.NewBook(TITLE, DESCRIPTION, ISBN, idGender, idAuthor);

            this.managementTest.ReadBook(0000);

            book = this.managementTest.FindBookByISBN(book.ISBN);

            #endregion
        }

        [TestMethod]
        [Description("Find all the books")]
        public void FindAllBooksTest()
        {
            #region Case 1: Get all genders

            List<Livro> expected = new List<Livro>(this.GenerateLivros(10));

            foreach (Livro livro in expected)
            {
                this.managementTest.NewBook(livro.Titulo, livro.Resumo, livro.ISBN, livro.IdGenero, livro.IdAutor);
            }

            List<Livro> actual = this.managementTest.FindAll();

            CollectionAssert.AreEqual(expected, actual);

            #endregion
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Generate Livro items 
        /// </summary>
        /// <param name="count">Number of items</param>
        /// <returns>List of Livro</returns>
        private IEnumerable<Livro> GenerateLivros(int count)
        {
            for (int i = 0; i < count; i++)
            {
                Livro book = new Livro();
                book.Titulo = String.Concat("Book title: ", i);
                book.Resumo = String.Concat("Book summary: ", i);
                book.IdAutor = this.idAuthor;
                book.IdGenero = this.idGender;
                book.ISBN = i;
                yield return book;
            }
        }

        #endregion
    }
}
