import os
from typing import List


def get_invalid_nicknames(path_base: str, nicknames: List[str]) -> List[str]:
    return list(filter(lambda nickname: not os.path.isfile(path_base + nickname + ".txt"), nicknames))
