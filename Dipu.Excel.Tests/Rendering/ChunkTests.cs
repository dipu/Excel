using System;
using System.Collections.Generic;
using Dipu.Excel.Embedded;
using Xunit;
using System.Linq;
using Moq;

namespace Dipu.Excel.Tests.Embedded
{
    public class ChunkTests
    {
        [Fact]
        public void OneChunk_OneRow_TwoCells()
        {
            var cells = new[]
            {
                new Cell(null, 0, 0),
                new Cell(null, 0, 1),
            };

            var chunks = cells.Chunks((v, w) => true).ToArray();
            Assert.Single(chunks);

            chunks = cells.Chunks((v, w) => false).ToArray();
            Assert.Equal(2, chunks.Length);
        }

        [Fact]
        public void OneChunk_OneRow_FourCells()
        {
            var cells = new[]
                {
                    new Cell(null, 0,0),
                    new Cell(null, 0,1),
                    new Cell(null, 0,2),
                    new Cell(null, 0,3),
            };

            var chunks = cells.Chunks((v, w) => true).ToArray();
            Assert.Single(chunks);
        }

        [Fact]
        public void OneChunk_TwoRows_OneCell()
        {
            var cells = new[]
                {
                    new Cell(null, 0,0),
                    new Cell(null, 1,0),
            };

            var chunks = cells.Chunks((v, w) => true).ToArray();
            Assert.Single(chunks);
        }


        [Fact]
        public void OneChunk_TwoRows_TwoCells()
        {
            var cells = new[]
                {
                    new Cell(null, 0,0),
                    new Cell(null,0,1),
                    new Cell(null,1,0),
                    new Cell(null,1,1),
            };

            var chunks = cells.Chunks((v, w) => true).ToArray();
            Assert.Single(chunks);
        }


        [Fact]
        public void TwoChunks_OneRow_TwoCells()
        {
            var cells = new[]
                {
                    new Cell(null, 0,0),
                    new Cell(null, 0,1),
                    new Cell(null, 0,3),
                    new Cell(null,0,4),
            };

            var chunks = cells.Chunks((v, w) => true).ToArray();
            Assert.Equal(2, chunks.Length);
        }


        [Fact]
        public void Square()
        {
            var raster = new[]
            {
                "###",
                "# #",
                "###",
            };

            var worksheet = new Mock<IEmbeddedWorksheet>().Object;
            var cells = CellsFromRaster(worksheet, raster, (v, c) => v.NumberFormat = c);

            var chunks = cells.Chunks((v, w) => Equals(v.NumberFormat, w.NumberFormat)).ToArray();

            Assert.Equal(5, chunks.Length);
        }

        [Fact]
        public void Cross()
        {
            var raster = new[]
            {
                "# #",
                " # ",
                "# #",
            };

            var worksheet = new Mock<IEmbeddedWorksheet>().Object;
            var cells = CellsFromRaster(worksheet, raster, (v, c) => v.NumberFormat = c);

            var chunks = cells.Chunks((v, w) => Equals(v.NumberFormat, w.NumberFormat)).ToArray();

            Assert.Equal(9, chunks.Length);
        }

        [Fact]
        public void HorizontalLines()
        {
            var raster = new[]
            {
                "###",
                "   ",
                "###",
            };

            var worksheet = new Mock<IEmbeddedWorksheet>().Object;
            var cells = CellsFromRaster(worksheet, raster, (v, c) => v.NumberFormat = c);

            var chunks = cells.Chunks((v, w) => Equals(v.NumberFormat, w.NumberFormat)).ToArray();

            Assert.Equal(3, chunks.Length);
        }

        [Fact]
        public void VerticalLines()
        {
            var raster = new[]
            {
                "# #",
                "# #",
                "# #",
            };

            var worksheet = new Mock<IEmbeddedWorksheet>().Object;
            var cells = CellsFromRaster(worksheet, raster, (v, c) => v.NumberFormat = c);

            var chunks = cells.Chunks((v, w) => Equals(v.NumberFormat, w.NumberFormat)).ToArray();

            Assert.Equal(3, chunks.Length);
        }

        [Fact]
        public void LShape()
        {
            var raster = new[]
            {
                "#  ",
                "#  ",
                "###",
            };

            var worksheet = new Mock<IEmbeddedWorksheet>().Object;
            var cells = CellsFromRaster(worksheet, raster, (v, c) => v.NumberFormat = c);

            var chunks = cells.Chunks((v, w) => Equals(v.NumberFormat, w.NumberFormat)).ToArray();

            Assert.Equal(3, chunks.Length);
        }

        [Fact]
        public void ReverseLShape()
        {
            var raster = new[]
            {
                "  #",
                "  #",
                "###",
            };

            var worksheet = new Mock<IEmbeddedWorksheet>().Object;
            var cells = CellsFromRaster(worksheet, raster, (v, c) => v.NumberFormat = c);

            var chunks = cells.Chunks((v, w) => Equals(v.NumberFormat, w.NumberFormat)).ToArray();

            Assert.Equal(3, chunks.Length);
        }

        private static IList<Cell> CellsFromRaster(IEmbeddedWorksheet worksheet, string[] raster, Action<ICell, string> setup)
        {
            var cells = new List<Cell>();
            for (var i = 0; i < raster.Length; i++)
            {
                var line = raster[i];
                for (var j = 0; j < 3; j++)
                {
                    var cell = new Cell(worksheet, i, j);
                    setup(cell, line[j].ToString());
                    cells.Add(cell);
                }
            }

            return cells;
        }
    }
}
