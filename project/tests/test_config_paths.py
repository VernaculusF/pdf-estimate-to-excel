import unittest
from pathlib import Path

import config


class ConfigPathTests(unittest.TestCase):
    def test_default_paths_point_to_workspace_level_folders(self):
        project_dir = Path(config.__file__).resolve().parent
        workspace_dir = project_dir.parent

        self.assertEqual(Path(config.DEFAULT_INPUT_DIR), workspace_dir / config.INPUT_DIR_NAME)
        self.assertEqual(Path(config.DEFAULT_OUTPUT_DIR), workspace_dir / config.OUTPUT_DIR_NAME)


if __name__ == "__main__":
    unittest.main()
