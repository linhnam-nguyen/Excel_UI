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

using BH.Adapter;
using BH.oM.Base;
using BH.Engine.Base;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;


namespace BH.UI.Excel
{
    public partial class AddIn : IExcelAddIn
    {
        /*******************************************/
        /**** Public Methods                    ****/
        /*******************************************/
        public static void SetAdapter(Range selection)
        {
            if(selection.Count != 1)
            {
                BH.Engine.Base.Compute.RecordError("Only one Adapter is accepted !");
                return;
            }

            object value = selection.Value;

            if (value == null)
            {
                m_Adapter = null;
                return;
            }

            object obj = GetObject(value as string);

            if (obj == null)
            {
                m_Adapter = null;
                return;
            }

            BHoMAdapter adapter = obj as BHoMAdapter;

            if (adapter != null) 
            {
                m_Adapter = adapter;
                m_AdapterName = selection.Value as string;
            }
            else
            {
                m_Adapter = null;
            }
        }

        /*******************************************/

        public static void SetAdapter(string adapterName)
        {
            if (string.IsNullOrEmpty(adapterName))
            {
                m_Adapter = null;
                m_AdapterName = string.Empty;
                return;
            }

            BHoMAdapter adapter = CreateAdapter(adapterName);
            if (adapter != null)
            {
                m_Adapter = adapter;
                m_AdapterName = adapterName;
                SetAdapterStatus(_isLive);
            }
            else
            {
                m_Adapter = null;
                m_AdapterName = string.Empty;
            }
        }

        /*******************************************/

        public static void SetAdapterStatus(bool isLive)
        {
            _isLive = isLive;
            if (m_AdapterName.Contains("On") || m_AdapterName.Contains("Off"))
            {
                if (_isLive)
                {
                    m_AdapterName = m_AdapterName.Replace(" (Off)", " (On)");
                }
                else
                {
                    m_AdapterName = m_AdapterName.Replace(" (On)", " (Off)");
                }
            }
            else
            {

                m_AdapterName = _isLive ? $"{m_AdapterName} (On)" : $"{m_AdapterName} (Off)";
            }
            string adapterName = m_AdapterName.Replace(" (Off)", "").Replace(" (On)", "");
            m_Adapter = CreateAdapter(adapterName);
        }

        /*******************************************/

        public static List<string> GetAvailableAdapters()
        {
            _adapters = new List<string>() { "None"};
            _adapters.AddRange(Query.AdapterTypeList()
                    .Where(t => t.IsClass && !t.IsAbstract && t.Name.EndsWith("Adapter"))
                    .Select(t => t.Name.Replace("Adapter", ""))
                    .Distinct()
                    .ToList());
            return _adapters;
        }

        /*******************************************/

        public static int GetSelectedAdapterIndex()
        {
            return _selectedAdapterIndex;
        }

        /*******************************************/

        public static string GetAdapterName()
        {
            if (m_Adapter != null)
            {
                return m_AdapterName;
            }
            else
            {
                return string.Empty;
            }
        }

        /*******************************************/

        private static BHoMAdapter CreateAdapter(string adapterName)
        {
            if (adapterName == "None")
            {
                return null;
            }

            string typeName = $"BH.Adapter.{adapterName}.{adapterName}Adapter";
            string settingsTypeName = $"BH.oM.Adapters.{adapterName}.Settings.{adapterName}Settings";
            try 
            {
            Type adapterType = Query.AdapterTypeList().FirstOrDefault(x => x.FullName == typeName);
            Type settingsType = BH.Engine.Base.Create.Type(settingsTypeName);  
                if (adapterType != null)
                {

                    dynamic settings = Activator.CreateInstance(settingsType);
                    dynamic adapter = Activator.CreateInstance(adapterType, settings, _isLive);
                    return adapter;
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        /*******************************************/
        /**** Private Fields                   *****/
        /*******************************************/
        private static BHoMAdapter m_Adapter;
        private static string m_AdapterName = string.Empty;
        private static List<string> _adapters = new List<string> ();
        private static int _selectedAdapterIndex = 0;
        private static bool _isLive = false;
    }
}






