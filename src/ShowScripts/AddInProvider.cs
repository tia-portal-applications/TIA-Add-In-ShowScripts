using Siemens.Engineering;
using Siemens.Engineering.AddIn;
using Siemens.Engineering.AddIn.Menu;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ShowScripts
{
    public sealed class AddInProvider : ProjectTreeAddInProvider //Definiert den Bereich, in dem das Add-In verfügbar sein soll
    {
        private readonly TiaPortal _tiaPortal;

        public AddInProvider(TiaPortal tiaPortal)
        {
            _tiaPortal = tiaPortal; //TIA-Portal-Instanz des ausgewählten AddIn
        }

        protected override IEnumerable<ContextMenuAddIn> GetContextMenuAddIns() //mein AddIn wird in Kontext angezeigt (Rechtsklick auf Screen)
        {
            yield return new AddIn(_tiaPortal); //erezugt ein neues AddIn meines Screens
        }
    }
}
