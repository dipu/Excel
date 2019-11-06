using System.Collections.Generic;
using System.Linq;
using Dipu.Excel.Rendering;
using Xunit;

namespace Dipu.Excel.Tests.Embedded
{
    public class BatchTests
    {
        [Fact]
        public void OneChunk_OneRow_TwoCells()
        {
            var cellValues = new[]
            {
                new CellValue(0, 0, "0,0"),
                new CellValue(0, 1, "0,1"),
            };

            var chunks = cellValues.Chunks().ToArray();
            Assert.Single(chunks);
        }

        [Fact]
        public void OneChunk_OneRow_FourCells()
        {
            var cellValues = new[]
                {
                    new CellValue(0,0,"0,0"),
                    new CellValue(0,1,"0,1"),
                    new CellValue(0,2,"0,2"),
                    new CellValue(0,3,"0,3"),
            };

            var chunks = cellValues.Chunks().ToArray();
            Assert.Single(chunks);
        }

        [Fact]
        public void OneChunk_TwoRows_OneCell()
        {
            var batch = new[]
                {
                    new CellValue(0,0,"0,0"),
                    new CellValue(1,0,"1,0"),
            };

            var chunks = batch.Chunks().ToArray();
            Assert.Single(chunks);
        }


        [Fact]
        public void OneChunk_TwoRows_TwoCells()
        {
            var cellValues = new[]
                {
                    new CellValue(0,0,"0,0"),
                    new CellValue(0,1,"0,1"),
                    new CellValue(1,0,"1,0"),
                    new CellValue(1,1,"1,1"),
            };

            var chunks = cellValues.Chunks().ToArray();
            Assert.Single(chunks);
        }


        [Fact]
        public void TwoChunks_OneRow_TwoCells()
        {
            var cellValues = new[]
                {
                    new CellValue(0,0,"0,0"),
                    new CellValue(0,1,"0,1"),
                    new CellValue(0,3,"0,3"),
                    new CellValue(0,4,"0,4"),
            };

            var chunks = cellValues.Chunks().ToArray();
            Assert.Equal(2, chunks.Length);
        }
    }
}
