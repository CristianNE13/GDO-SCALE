using System;
using System.Collections.Generic;
using System.Linq;

namespace Scale_Program.Functions
{
    public sealed class ComboResult
    {
        public int[] ArticuloIds { get; set; }       
        public decimal PesoMin { get; set; }         
        public decimal PesoMax { get; set; }          
        public List<ArticleActivation> Articulos { get; set; } 
        public string PasosCsv { get; set; }                   
    }

    public sealed class ArticleActivation
    {
        public int ArticuloID { get; set; }
        public string NoParte { get; set; }          
        public decimal PesoMin { get; set; }         
        public decimal PesoMax { get; set; }          
        public int Paso { get; set; }                 
    }

     public sealed class CombosProcesoResult
    {
        public List<ComboResult> Combos { get; set; }
        public List<Articulo> Ccam { get; set; }

        public CombosProcesoResult()
        {
            Combos = new List<ComboResult>();
            Ccam = new List<Articulo>();
        }
    }

    public sealed class ComboCalculator
    {
        public const decimal EPS = 0.0005m;

        private static IEnumerable<int[]> GetAllCombinations(IReadOnlyList<int> items)
        {
            int n = items.Count;
            if (n == 0) yield break;
            int total = 1 << n;
            for (int mask = 1; mask < total; mask++)
            {
                List<int> combo = new List<int>(n);
                for (int i = 0; i < n; i++)
                    if ((mask & (1 << i)) != 0)
                        combo.Add(items[i]);
                yield return combo.ToArray();
            }
        }

        private static bool EsCCAM(string noParte)
        {
            return !string.IsNullOrWhiteSpace(noParte) &&
                   noParte.IndexOf("CCAM", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public CombosProcesoResult BuildCombosDelProceso(int modproceso, int proceso)
        {
            CombosProcesoResult result = new CombosProcesoResult();

            using (dc_missingpartsEntities db = new dc_missingpartsEntities())
            {
                var items = db.Articulos
                    .Where(x => x.ModProceso == modproceso && x.Proceso == proceso)
                    .Select(x => new
                    {
                        x.Id,
                        x.Cantidad,
                        x.Paso,
                        x.NoParte,
                        x.PesoMin,
                        x.PesoMax
                    })
                    .ToList();

                result.Ccam = items
                    .Where(x => EsCCAM(x.NoParte))
                    .Select(x => new Articulo
                    {
                        Id = x.Id,
                        Cantidad = x.Cantidad,
                        Proceso = proceso,
                        ModProceso = modproceso,
                        Paso = x.Paso
                    })
                    .ToList();

                var calcItems = items.Where(x => !EsCCAM(x.NoParte)).ToList();
                if (calcItems.Count == 0)
                    return result;

                var ids = calcItems.Select(x => x.Id).ToList();

                var byId = calcItems.ToDictionary(
                    k => k.Id,
                    v => new
                    {
                        v.NoParte,
                        v.Paso,
                        v.Cantidad,
                        v.PesoMin,
                        v.PesoMax
                    });

                foreach (var combo in GetAllCombinations(ids))
                {
                    decimal totalMin = 0m, totalMax = 0m;
                    List<ArticleActivation> articulos = new List<ArticleActivation>(combo.Length);
                    HashSet<int> pasos = new HashSet<int>();

                    foreach (int id in combo)
                    {
                        var d = byId[id];
                        decimal artMin = (decimal)(d.PesoMin * d.Cantidad);
                        decimal artMax = (decimal)(d.PesoMax * d.Cantidad);

                        articulos.Add(new ArticleActivation
                        {
                            ArticuloID = id,
                            NoParte = d.NoParte,
                            Paso = d.Paso,
                            PesoMin = artMin,
                            PesoMax = artMax
                        });

                        pasos.Add(d.Paso);

                        totalMin += artMin;
                        totalMax += artMax;
                    }

                    string pasosCsv = string.Join(",",
                        pasos.OrderBy(p => p).Select(p => p.ToString()));

                    result.Combos.Add(new ComboResult
                    {
                        ArticuloIds = combo,
                        PesoMin = totalMin,
                        PesoMax = totalMax,
                        Articulos = articulos,
                        PasosCsv = pasosCsv
                    });
                }
            }

            return result;
        }

        public List<ComboResult> BuscarCombosQueEmpatanPeso(
            IEnumerable<ComboResult> combos, decimal peso, decimal eps)
        {
            return combos
                .Where(c => c.PesoMin - eps <= peso && peso <= c.PesoMax + eps)
                .OrderBy(c => c.PesoMax - c.PesoMin)
                .ToList();
        }

        public List<ComboResult> BuscarCombosQueEmpatanPeso(
            IEnumerable<ComboResult> combos, decimal peso)
        {
            return BuscarCombosQueEmpatanPeso(combos, peso, EPS);
        }
    }
}
