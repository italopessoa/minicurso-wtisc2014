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
    /// Class to manager information of the Gender
    /// </summary>
    public class MGenero
    {
        BooksEntities entities;

        /// <summary>
        /// 
        /// </summary>
        public MGenero()
        {
            this.entities = new BooksEntities();
        }

        /// <summary>
        /// Create a new Gender
        /// </summary>
        /// <param name="name">The name of the gender</param>
        /// <param name="description">The description of the gender</param>
        /// <exception cref="WTISC._2014.Data.Exceptions.AuthorException"></exception>
        /// <returns>New Gender</returns>
        public Genero NewGender(string name, string description)
        {
            if (!string.IsNullOrEmpty(name))
            {
                if (this.FindGenderByName(name) != null)
                {
                    Genero genero = new Genero();
                    genero.Nome = "Novo Gênero";
                    genero.Descricao = "Descrição de um novo gênero";

                    this.entities.Genero.Add(genero);
                    this.entities.SaveChanges();

                    return genero;
                }
                else
                {
                    throw new GenderException("The name can't be empty!");
                }
            }
            else
            {
                throw new GenderException("The name can't be empty!");
            }
        }

        /// <summary>
        /// Find a gender by name
        /// </summary>
        /// <param name="name">The name of the gender</param>
        /// <returns></returns>
        public Genero FindGenderByName(string name)
        {
            return this.entities.Genero.FirstOrDefault<Genero>(g => g.Nome.Equals(name));
        }

        /// <summary>
        /// Returns all Genders in the database
        /// </summary>
        /// <returns>List of genders</returns>
        public List<Genero> FindAll()
        {
            return this.entities.Genero.ToList<Genero>();
        }
    }
}
