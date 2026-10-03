"""Automated unit test suite for B2B Web Scraper & ETL Pipeline."""
import unittest
from pathlib import Path
from bs4 import BeautifulSoup
import openpyxl
from scraper import B2BDataScraper

class TestB2BDataScraper(unittest.TestCase):
    def setUp(self):
        self.scraper = B2BDataScraper()

    def test_scraper_initialization(self):
        """Verify scraper initializes with correct default configuration."""
        self.assertIsNotNone(self.scraper.session)
        self.assertTrue(self.scraper.base_url.startswith("http"))
        self.assertEqual(len(self.scraper.results), 0)

    def test_get_headers_structure(self):
        """Verify headers contain realistic User-Agent and Accept values."""
        headers = self.scraper._get_headers()
        self.assertIn("User-Agent", headers)
        self.assertIn("Mozilla", headers["User-Agent"])
        self.assertIn("Accept", headers)

    def test_parse_product_item_valid_html(self):
        """Verify product item HTML is correctly parsed into structured fields."""
        html = """
        <article class="product_pod">
            <div class="image_container">
                <a href="catalogue/a-light-in-the-attic_1000/index.html">
                    <img class="thumbnail" src="media/cache/test.jpg" alt="A Light in the Attic"/>
                </a>
            </div>
            <p class="star-rating Three">
                <i class="icon-star"></i>
            </p>
            <h3><a href="catalogue/a-light-in-the-attic_1000/index.html" title="A Light in the Attic">A Light in the Attic</a></h3>
            <div class="product_price">
                <p class="price_color">£51.77</p>
                <p class="instock availability"><i class="icon-ok"></i> In stock</p>
            </div>
        </article>
        """
        soup = BeautifulSoup(html, "html.parser")
        article = soup.find("article")
        parsed = self.scraper.parse_product_item(article, "http://books.toscrape.com/index.html")
        
        self.assertIsNotNone(parsed)
        self.assertEqual(parsed["Tên sản phẩm"], "A Light in the Attic")
        self.assertEqual(parsed["Đánh giá (Sao)"], 3)
        self.assertEqual(parsed["Tình trạng tồn kho"], "Còn hàng")
        self.assertGreater(parsed["Giá niêm yết (GBP)"], 0)

    def test_sample_output_csv_exists_and_valid(self):
        """Verify pre-generated sample CSV output file exists and has rows."""
        csv_path = Path(__file__).resolve().parent.parent / "sample_output.csv"
        self.assertTrue(csv_path.exists(), "sample_output.csv should exist")
        with open(csv_path, "r", encoding="utf-8") as f:
            lines = [l.strip() for l in f if l.strip()]
        self.assertGreater(len(lines), 1, "CSV should contain header plus data rows")

    def test_sample_output_xlsx_exists_and_valid(self):
        """Verify pre-generated sample XLSX output exists and has valid sheet structure."""
        xlsx_path = Path(__file__).resolve().parent.parent / "sample_output.xlsx"
        self.assertTrue(xlsx_path.exists(), "sample_output.xlsx should exist")
        wb = openpyxl.load_workbook(xlsx_path)
        sheet = wb.active
        self.assertGreater(sheet.max_row, 1, "Excel sheet should have data rows")

if __name__ == "__main__":
    unittest.main()
