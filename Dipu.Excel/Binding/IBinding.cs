// <copyright file="Client.cs" company="Allors bvba">
// Copyright (c) Allors bvba. All rights reserved.
// Licensed under the LGPL license. See LICENSE file in the project root for full license information.
// </copyright>

namespace Dipu.Excel
{
    public interface IBinding
    {
        void ToCell(ICell cell);

        void ToDomain(ICell cell);
    }
}