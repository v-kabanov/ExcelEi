using System;
using NUnit.Framework;
using ExcelEi.Read;

namespace ExcelEi.Test
{
    [TestFixture]
    public class ConversionTest
    {
        [Test]
        public void TextToInt()
        {
            var result = Conversion.GetTypedExcelValue<int>("204");

            Assert.AreEqual(204, result);
        }

        [Test]
        public void DoubleToNullableInt()
        {
            var result = Conversion.GetTypedExcelValue<int?>(2D);

            Assert.AreEqual(2, result);
        }

        [Test]
        public void StringToDecimal()
        {
            var result = Conversion.GetTypedExcelValue<decimal>("1.4");

            Assert.AreEqual(1.4, result);
        }

        [Test]
        public void EmptyStringToNullableDecimal()
        {
            var result = Conversion.GetTypedExcelValue<decimal?>("");

            Assert.AreEqual(null, result);
        }

        [Test]
        public void BlankStringToNullableDecimal()
        {
            var result = Conversion.GetTypedExcelValue<decimal?>(" ");

            Assert.AreEqual(null, result);
        }

        [Test]
        public void EmptyStringToDecimal()
        {
            TestDelegate testDelegate = () => Conversion.GetTypedExcelValue<decimal>("");
            Assert.That(testDelegate, Throws.InstanceOf<FormatException>());
        }

        [Test]
        public void FloatingPointStringToInt()
        {
            TestDelegate testDelegate = () => Conversion.GetTypedExcelValue<int>("1.4");
            Assert.That(testDelegate, Throws.InstanceOf<FormatException>());
        }

        [Test]
        public void IntToDateTime()
        {
            Assert.Throws<InvalidCastException>(() => Conversion.GetTypedExcelValue<DateTime>(122));
        }

        [Test]
        public void IntToTimeSpan()
        {
            Assert.Throws<InvalidCastException>(() => Conversion.GetTypedExcelValue<TimeSpan>(122));
        }

        [Test]
        public void IntStringToTimeSpan()
        {
            Assert.AreEqual(TimeSpan.FromDays(122), Conversion.GetTypedExcelValue<TimeSpan>("122"));
        }

        [Test]
        public void BoolToInt()
        {
            Assert.AreEqual(1, Conversion.GetTypedExcelValue<int>(true));
            Assert.AreEqual(0, Conversion.GetTypedExcelValue<int>(false));
        }

        [Test]
        public void BoolToDecimal()
        {
            Assert.AreEqual(1m, Conversion.GetTypedExcelValue<decimal>(true));
            Assert.AreEqual(0m, Conversion.GetTypedExcelValue<decimal>(false));
        }

        [Test]
        public void BoolToDouble()
        {
            Assert.AreEqual(1d, Conversion.GetTypedExcelValue<double>(true));
            Assert.AreEqual(0d, Conversion.GetTypedExcelValue<double>(false));
        }

        [Test]
        public void BadTextToInt()
        {
            Assert.Throws<FormatException>(() => Conversion.GetTypedExcelValue<int>("text1"));
        }
    }
}
