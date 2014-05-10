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

namespace WTISC._2014.Data.Test
{
    [TestClass]
    public class AutorTest
    {
        #region Private Members

        private MAutor managementTest;
        private const string NAME = "author test";

        #endregion

        #region Autor Tests

        [TestInitialize]
        [Description("Initialize test")]
        public void Initialize()
        {
            this.managementTest = new MAutor();
        }

        [TestCleanup]
        [Description("Delete database values")]
        public void RestartDB()
        {
            BooksEntities resetEntities = new BooksEntities();
            resetEntities.Database.ExecuteSqlCommand("DELETE [Livro]");
            resetEntities.Database.ExecuteSqlCommand("DELETE [Autor]");
        }

        [TestMethod]
        [Description("Create an author")]
        public void NewAuthorTest()
        {
            #region Case 1: Create a new author

            Autor authorTest = this.managementTest.NewAuthor(NAME);
            Assert.AreEqual(NAME, authorTest.Nome);

            #endregion
        }

        [TestMethod]
        [ExpectedException(typeof(AuthorException), "The author already exists!")]
        [Description("Create an author with already existing name")]
        public void NewAutorException_NameTest()
        {
            #region Case 1: Create a author with already existing name

            this.managementTest.NewAuthor(NAME);

            this.managementTest.NewAuthor(NAME);

            #endregion
        }

        [TestMethod]
        [ExpectedException(typeof(AuthorException), "The name can't be empty!")]
        [Description("Create a author without name")]
        public void NewAutorException_EmptyNameTest()
        {
            #region Case 1: Create a author without name

            this.managementTest.NewAuthor(String.Empty);

            #endregion
        }

        [TestMethod]
        [Description("Find author by name")]
        public void FindAuthorByNameTest()
        {
            #region Case 1: Get a created author

            Autor expected = this.managementTest.NewAuthor(NAME);
            Autor actual = this.managementTest.FindAuthorByName(NAME);

            Assert.AreEqual(expected, actual);

            #endregion

            #region Case 2: No result

            expected = null;
            actual = this.managementTest.FindAuthorByName("invalid name");

            Assert.IsNull(actual);

            #endregion
        }

        [TestMethod]
        [Description("Alter author's name")]
        public void UpdateAuthorTest()
        {
            #region Case 1: Update the author's informations

            Autor expected = this.managementTest.NewAuthor(NAME);
            expected.Nome = "altered name";

            this.managementTest.UpdateAuthor(expected);

            Autor actual = this.managementTest.FindAuthorByName(expected.Nome);

            Assert.AreEqual(expected.Nome, actual.Nome);
            Assert.AreEqual(expected, actual);

            #endregion
        }

        [TestMethod]
        [ExpectedException(typeof(AuthorException), "The author already exists!")]
        [Description("Alter the author's name")]
        public void UpdateAuthorExceptionTest()
        {
            #region Case 1: Alter author's name for already existing name

            Autor author = this.managementTest.NewAuthor(NAME);
            this.managementTest.UpdateAuthor(author);

            #endregion
        }

        [TestMethod]
        [Description("Alter the author's name with an already existing name")]
        public void FindAllAuthorsTest()
        {
            #region Case 1: Get all authors

            List<Autor> expected = new List<Autor>(this.GenerateAuthors(10));

            foreach (Autor author in expected)
            {
                author.Id = this.managementTest.NewAuthor(author.Nome).Id;
            }

            List<Autor> actual = this.managementTest.FindAll();

            CollectionAssert.AreEqual(expected, actual);

            #endregion
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Generate Autor items 
        /// </summary>
        /// <param name="count">Number of items</param>
        /// <returns>List of Autor</returns>
        private IEnumerable<Autor> GenerateAuthors(int count)
        {
            for (int i = 0; i < count; i++)
            {
                Autor author = new Autor();
                author.Nome = String.Concat("AuthorName: ", i);
                yield return author;
            }
        }

        #endregion
    }
}
