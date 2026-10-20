import unittest
from src.updater import parse_version_tuple

class TestUpdater(unittest.TestCase):
    def test_parse_version_tuple(self):
        self.assertEqual(parse_version_tuple("v1.2.15"), (1, 2, 15))
        self.assertEqual(parse_version_tuple("1.2.0"), (1, 2, 0))
        self.assertEqual(parse_version_tuple("v2026.10.20"), (2026, 10, 20))
        self.assertEqual(parse_version_tuple(""), (0, 0, 0))
        
        # Kiểm tra so sánh phiên bản
        v1 = parse_version_tuple("v1.2.14")
        v2 = parse_version_tuple("v1.2.15")
        v3 = parse_version_tuple("v1.3.0")
        self.assertLess(v1, v2)
        self.assertLess(v2, v3)
        self.assertEqual(v1, parse_version_tuple("1.2.14"))

if __name__ == "__main__":
    unittest.main()
