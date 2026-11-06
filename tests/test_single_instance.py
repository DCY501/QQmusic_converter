import os
import subprocess
import sys
import unittest
import uuid
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from single_instance import SingleInstance


@unittest.skipUnless(os.name == 'nt', 'Windows kernel objects required')
class SingleInstanceTests(unittest.TestCase):
    def test_duplicate_is_blocked_and_activation_is_delivered(self):
        name = 'Local\\QQMusicTest-' + uuid.uuid4().hex
        first, second = SingleInstance(name), SingleInstance(name)
        self.addCleanup(first.close)
        self.addCleanup(second.close)
        self.assertTrue(first.acquire())
        self.assertFalse(second.acquire())
        self.assertTrue(second.notify_existing())
        self.assertTrue(first.activation_requested())
        self.assertFalse(first.activation_requested())
        first.close()
        self.assertTrue(second.acquire())

    def test_lock_is_released_after_owner_process_is_killed(self):
        name = 'Local\\QQMusicTest-' + uuid.uuid4().hex
        code = ('import sys; from single_instance import SingleInstance; '
                'guard=SingleInstance(sys.argv[1]); print(guard.acquire(), flush=True); '
                'sys.stdin.readline()')
        child = subprocess.Popen([sys.executable, '-u', '-c', code, name],
                                 cwd=Path(__file__).resolve().parents[1] / 'src',
                                 stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        guard = SingleInstance(name)
        try:
            self.assertEqual(child.stdout.readline().strip(), 'True')
            self.assertFalse(guard.acquire())
            child.terminate()
            child.communicate(timeout=5)
            self.assertTrue(guard.acquire())
        finally:
            guard.close()
            if child.poll() is None:
                child.kill()
            child.communicate(timeout=5)


if __name__ == '__main__':
    unittest.main()
