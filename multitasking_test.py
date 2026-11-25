from task_manager import TaskManager
import time


def task_one():
    time.sleep(3)
    return "Task 1: Research preparation completed."


def task_two():
    time.sleep(2)
    return "Task 2: Calculation completed."


def task_three():
    time.sleep(4)
    return "Task 3: File preparation completed."


def show_result(task_id, status, result):
    print(f"\n[{task_id}] {status}: {result}")


manager = TaskManager(max_workers=3)

print("A.L.P.H.A - 6 MULTITASKING TEST")
print("------------------------")

manager.submit(
    "Research preparation",
    task_one,
    on_complete=show_result
)

manager.submit(
    "Calculation",
    task_two,
    on_complete=show_result
)

manager.submit(
    "File preparation",
    task_three,
    on_complete=show_result
)

print("\nAll three tasks have started.")
print("A.L.P.H.A - 6 can continue accepting other work.")

try:
    while True:
        tasks = manager.get_tasks()

        if all(
            task["status"] in ("Completed", "Failed")
            for task in tasks.values()
        ):
            break

        time.sleep(0.5)

except KeyboardInterrupt:
    print("\nTest interrupted.")

finally:
    manager.shutdown()

print("\nFinal task status:")
for task_id, task in manager.get_tasks().items():
    print(task_id, task["name"], "-", task["status"])