//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace WTISC._2014.Data
{
    using System;
    using System.Collections.Generic;
    
    public partial class Autor
    {
        public Autor()
        {
            this.Livro = new HashSet<Livro>();
        }
    
        public int Id { get; set; }
        public string Nome { get; set; }
    
        public virtual ICollection<Livro> Livro { get; set; }
    }
}
