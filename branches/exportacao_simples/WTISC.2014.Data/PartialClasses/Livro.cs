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

namespace WTISC._2014.Data
{
    /// <summary>
    /// Partial class to manual code
    /// </summary>
    public partial class Livro
    {
        /// <summary>
        /// Compare instances
        /// </summary>
        /// <param name="obj">Instance to compare</param>
        /// <returns>True or False</returns>
        public override bool Equals(object obj)
        {
            if (obj == null)
            {
                return false;
            }
            else if (obj is Livro)
            {
                Livro aux = obj as Livro;
                return (
                    this.ISBN == aux.ISBN
                    && this.IdGenero == aux.IdGenero
                    && this.IdAutor == aux.IdAutor
                    && this.Titulo.Equals(aux.Titulo)
                    && this.Resumo.Equals(aux.Resumo)
                    );
            }
            else
            {
                return false;
            }
        }
    }
}
