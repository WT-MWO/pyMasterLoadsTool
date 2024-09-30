import unittest
from pymasterloadstool import utilities


class TestUtilities(unittest.TestCase):

    def test_list_to_str(self):
        test_list = ["4", "5", "acv", "yts", 6, 7.0]
        desired_output = "4, 5, acv, yts, 6, 7.0"
        output = utilities.list_to_str(test_list)
        self.assertEqual(output, desired_output)

    def test_get_key(self):
        test_dictionary = {"a": "100", "b": 101, 3: 45}
        self.assertEqual(utilities.get_key(test_dictionary, "100"), "a")
        self.assertEqual(utilities.get_key(test_dictionary, 101), "b")
        self.assertEqual(utilities.get_key(test_dictionary, 45), 3)
