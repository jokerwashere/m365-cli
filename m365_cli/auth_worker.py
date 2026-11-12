"""Internal entry point: one authentication operation, never a service mutation."""

from .cli import main

if __name__ == "__main__":
    main(_auth_worker=True)
