from concurrent.futures import ThreadPoolExecutor
from threading import Lock
from datetime import datetime
import uuid


class TaskManager:
    def __init__(self, max_workers=3):
        self.executor = ThreadPoolExecutor(
            max_workers=max_workers,
            thread_name_prefix="A.L.P.H.A - 6-Task"
        )
        self.tasks = {}
        self.lock = Lock()
        self.closed = False

    def submit(self, task_name, function, *args, on_complete=None, **kwargs):
        """Run a function in the background."""

        with self.lock:
            if self.closed:
                raise RuntimeError("Task manager is closed.")

            task_id = str(uuid.uuid4())[:8]

            self.tasks[task_id] = {
                "name": task_name,
                "status": "Running",
                "result": None,
                "error": None,
                "created": datetime.now().strftime("%H:%M:%S")
            }

        future = self.executor.submit(function, *args, **kwargs)

        def task_finished(completed_future):
            try:
                result = completed_future.result()

                with self.lock:
                    self.tasks[task_id]["status"] = "Completed"
                    self.tasks[task_id]["result"] = result

                if on_complete:
                    on_complete(task_id, "Completed", result)

            except Exception as error:
                with self.lock:
                    self.tasks[task_id]["status"] = "Failed"
                    self.tasks[task_id]["error"] = str(error)

                if on_complete:
                    on_complete(task_id, "Failed", str(error))

        future.add_done_callback(task_finished)
        return task_id

    def get_tasks(self):
        """Return a safe copy of task information."""
        with self.lock:
            return {
                task_id: details.copy()
                for task_id, details in self.tasks.items()
            }

    def shutdown(self, wait=True):
        """Stop accepting tasks and finish existing tasks."""
        with self.lock:
            self.closed = True

        self.executor.shutdown(
            wait=wait,
            cancel_futures=False
        )