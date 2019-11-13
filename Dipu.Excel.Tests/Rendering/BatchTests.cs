using Dipu.Excel.Embedded;
using Xunit;
using System.Linq;

namespace Dipu.Excel.Tests.Embedded
{
    public class BatchTests
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
    }
}
