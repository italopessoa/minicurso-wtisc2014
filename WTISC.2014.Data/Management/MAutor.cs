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
    /// Class to manager information of the Author
    /// </summary>
    public class MAutor
    {
        BooksEntities entities;

        /// <summary>
        /// 
        /// </summary>
        public MAutor(){
            this.entities = new BooksEntities();
        }

        /// <summary>
        /// Create a new Author
        /// </summary>
        /// <param name="name">The name of the Book's Author</param>
        /// <exception cref="WTISC._2014.Data.Exceptions.AuthorException"></exception>
        /// <returns>New Author</returns>
        public Autor NewAuthor(string name)
        {
            if (!string.IsNullOrEmpty(name))
            {
                if (this.FindAuthorByName(name) == null)
                {
                    Autor autor = new Autor();
                    autor.Nome = name;
                    this.entities.Autor.Add(autor);
                    this.entities.SaveChanges();

                    return autor;
                }
                else
                {
                    throw new AuthorException("The author already exists!");
                }
            }
            else throw new AuthorException("The name can't be empty!");
        }

        /// <summary>
        /// Find Author by name
        /// </summary>
        /// <param name="name">The name of the Book's Author</param>
        /// <returns>Author</returns>
        public Autor FindAuthorByName(string name)
        {
            return this.entities.Autor.FirstOrDefault<Autor>(a => a.Nome.Equals(name));
        }

        /// <summary>
        /// Update Author data
        /// </summary>
        /// <param name="newAuthor">New values</param>
        /// <exception cref="WTISC._2014.Data.Exceptions.AuthorException"></exception>
        public void UpdateAuthor(Autor newAuthor)
        {
            if (this.FindAuthorByName(newAuthor.Nome) == null)
            {
                Autor author = entities.Autor.FirstOrDefault<Autor>(a => a.Id == newAuthor.Id);
                author.Nome = newAuthor.Nome;
                entities.SaveChanges();
            }
            else
            {
                throw new AuthorException("The author already exists!");
            }
        }
    
        /// <summary>
        /// Return all the authors
        /// </summary>
        /// <returns>List of authors</returns>
        public List<Autor> FindAll()
        {
            return this.entities.Autor.ToList<Autor>();
        }
    }
}
