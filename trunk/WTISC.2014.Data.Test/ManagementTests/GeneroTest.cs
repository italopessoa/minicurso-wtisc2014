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
    public class GeneroTest
    {
        private MGenero managementTest;
        const string NAME = "gender test";
        const string DESCRIPTION = "gender description test";

        [TestInitialize]
        public void Initialize()
        {
            this.managementTest = new MGenero();
        }

        [TestCleanup]
        public void RestartDB()
        {
            BooksEntities resetEntities = new BooksEntities();
            resetEntities.Database.ExecuteSqlCommand("DELETE [Livro]");
            resetEntities.Database.ExecuteSqlCommand("DELETE [Genero]");
        }

        [TestMethod]
        public void NewGenderTest()
        {
            #region Case 1: Create a new Gender

            Genero genderTest = this.managementTest.NewGender(NAME,DESCRIPTION);
            Assert.AreEqual(NAME, genderTest.Nome);

            #endregion
        }

        [TestMethod]
        [ExpectedException(typeof(GenderException), "The gender already exists!")]
        public void NewGenderException_NameTest()
        {
            #region Case 1: Create a gender with already existing name

            this.managementTest.NewGender(NAME, DESCRIPTION);
            this.managementTest.NewGender(NAME, "other description");

            #endregion
        }

        [TestMethod]
        [ExpectedException(typeof(GenderException), "The name can't be empty!")]
        public void NewGenderException_EmptyNameTest()
        {
            #region Case 1: Create a gender without name

            this.managementTest.NewGender(String.Empty,DESCRIPTION);

            #endregion
        }

        [TestMethod]
        public void FindGenderByNameTest()
        {
            #region Case 1: Get a created gender

            Genero expected = this.managementTest.NewGender(NAME,DESCRIPTION);
            Genero actual = this.managementTest.FindGenderByName(NAME);

            Assert.AreEqual(expected, actual);

            #endregion

            #region Case 2: No result

            expected = null;
            actual = this.managementTest.FindGenderByName("invalid name");

            Assert.AreEqual(expected, actual);

            #endregion
        }

        [TestMethod]
        public void FindAllGendersTest()
        {
            #region Case 1: Get all genders

            List<Genero> expected = new List<Genero>(this.GenerateGenders(10));

            foreach (Genero gender in expected)
            {
                gender.Id = this.managementTest.NewGender(gender.Nome,gender.Descricao).Id;
            }

            List<Genero> actual = this.managementTest.FindAll();

            CollectionAssert.AreEqual(expected, actual);

            #endregion
        }

        #region Private Methods

        /// <summary>
        /// Generate Genero items 
        /// </summary>
        /// <param name="count">Number of items</param>
        /// <returns>List of Genero</returns>
        private IEnumerable<Genero> GenerateGenders(int count)
        {
            for (int i = 0; i < count; i++)
            {
                Genero gender = new Genero();
                gender.Nome = String.Concat("GenderName: ", i);
                gender.Descricao = String.Concat("GenderDescription: ", i);
                yield return gender;
            }
        }

        #endregion
    }
}
