//    This file is part of QTTabBar, a shell extension for Microsoft
//    Windows Explorer.
//    Copyright (C) 2007-2010  Quizo, Paul Accisano
//
//    QTTabBar is free software: you can redistribute it and/or modify
//    it under the terms of the GNU General Public License as published by
//    the Free Software Foundation, either version 3 of the License, or
//    (at your option) any later version.
//
//    QTTabBar is distributed in the hope that it will be useful,
//    but WITHOUT ANY WARRANTY; without even the implied warranty of
//    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
//    GNU General Public License for more details.
//
//    You should have received a copy of the GNU General Public License
//    along with QTTabBar.  If not, see <http://www.gnu.org/licenses/>.

namespace QTPlugin {
    using System;

    public interface ITab {
        bool Browse(QTPlugin.Address address);
        bool Browse(bool fBack);
        void Clone(int index, bool fSelect);
        bool Close();
        QTPlugin.Address[] GetBraches();
        QTPlugin.Address[] GetHistory(bool fBack);
        bool Insert(int index);

        QTPlugin.Address Address { get; }

        int Index { get; }

        bool Locked { get; set; }

        bool Selected { get; set; }

        string SubText { get; set; }

        string Text { get; set; }
    }
}