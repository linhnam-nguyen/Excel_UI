/*
 * This file is part of the Buildings and Habitats object Model (BHoM)
 * Copyright (c) 2015 - 2025, the respective contributors. All rights reserved.
 *
 * Each contributor holds copyright over their respective contributions.
 * The project versioning (Git) records all such contribution source information.
 *
 *
 * The BHoM is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Lesser General Public License as published by
 * the Free Software Foundation, either version 3.0 of the License, or
 * (at your option) any later version.
 *
 * The BHoM is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public License
 * along with this code. If not, see <https://www.gnu.org/licenses/lgpl-3.0.html>.
 */

using BH.UI.Excel.Templates;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace BH.UI.Excel.Addin
{
    public partial class Ribbon : ExcelRibbon
    {
        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/

        public void SetAdapter(IRibbonControl control)
        {
            Application app = ExcelDnaUtil.Application as Application;
            AddIn.SetAdapter(app.Selection as Range);
            _ribbon.InvalidateControl("adapterName");
            _ribbon.InvalidateControl("adapterSelector");
        }

        public string GetAdapterName(IRibbonControl control)
        {
            return AddIn.GetAdapterName();
        }

        /*******************************************/
        /**** DropDown Callbacks                ****/
        /*******************************************/

        public int GetAdapterCount(IRibbonControl control)
        {
            _adapters = AddIn.GetAvailableAdapters();
            return _adapters.Count;
        }

        public string GetAdapterLabel(IRibbonControl control, int index)
        {
            return _adapters[index];
        }

        public int GetSelectedAdapterIndex(IRibbonControl control)
        {
            return AddIn.GetSelectedAdapterIndex();
        }

        public void OnAdapterChange(IRibbonControl control, string id, int index)
        {
            AddIn.SetAdapter(_adapters[index]);
            _adapters = AddIn.GetAvailableAdapters();
            _ribbon.InvalidateControl("adapterName");
        }

        /*******************************************/

        public void OnToggleConnect(IRibbonControl control, bool pressed)
        {
            this._isLive = pressed;
            AddIn.SetAdapterStatus(pressed);
            _adapters = AddIn.GetAvailableAdapters();
            _ribbon.InvalidateControl("adapterName");
        }

        public bool GetConnectState(IRibbonControl control)
        {
            return this._isLive;
        }

        /*******************************************/
        /**** Private Fields                   *****/
        /*******************************************/
        private List<string> _adapters = new List<string>();
        private bool _isLive = false;
    }
}