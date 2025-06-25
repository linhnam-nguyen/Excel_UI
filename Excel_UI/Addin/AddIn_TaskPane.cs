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

using BH.UI.Excel.Addin;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using System;
using System.Windows.Forms;
using System.Windows.Forms.Integration;


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
                System.Windows.Controls.UserControl wpfControl = new ExplorePanel();

                var host = new ElementHost
                {
                    Child = wpfControl,
                    Dock = DockStyle.Fill
                };

                var panel = new System.Windows.Forms.UserControl();
                panel.Controls.Add(host);


                try
                {
                    TaskPane = CustomTaskPaneFactory.CreateCustomTaskPane(panel, "BHoM Explorer");
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show($"Error creating task pane: {ex.Message}");
                    return;
                }
            }

            TaskPane.Visible = show;
        }
    }
}






