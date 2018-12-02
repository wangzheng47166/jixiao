from rule_readers.rule_reader import RulesReader
from rule_readers.user_map_reader import UserMapReader
from rule_writers.rule_writer import RuleWriter


def main():
    usermaprader = UserMapReader()
    usermaprader.read()

    rulereader = RulesReader()
    rulereader.read()

    ruleWriter = RuleWriter()
    ruleWriter.write()

main()
