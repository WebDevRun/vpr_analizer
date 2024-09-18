from argparse import ArgumentParser


class ArgParser:
    def __init__(self):
        self.parser = ArgumentParser()
        self.parser.add_argument(
            "--paint",
            dest="paint",
            action="store_true",
            help="add conditional formatting in the tables",
        )

    def get_args(self):
        args = self.parser.parse_args()
        return args
