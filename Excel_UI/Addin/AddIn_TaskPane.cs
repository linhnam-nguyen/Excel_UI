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

using System;
using System.IO;
using System.Reflection;
using System.Linq;
using ExcelDna.Integration;
using System.Collections.Generic;
using System.Collections;
using BH.oM.Base;
using ExcelDna.Integration.CustomUI;
using System.Windows.Forms.Integration;
using BH.UI.Excel.Addin;

namespace BH.UI.Excel
{
    public partial class AddIn : IExcelAddIn
    {
        /*******************************************/
        /**** Methods                           ****/
        /*******************************************/
        public static CustomTaskPane TaskPane { get; private set; } = null;

        public static void ToggleTaskPane(bool show)
        {
            if (TaskPane == null)
            {
                var wpfControl = new ExplorePanel();
                var host = new ElementHost
                {
                    Child = wpfControl,
                    Dock = System.Windows.Forms.DockStyle.Fill
                };

                TaskPane = CustomTaskPaneFactory.CreateCustomTaskPane(host, "My WPF Task Pane");
            }

            TaskPane.Visible = show;
        }
    }
}






