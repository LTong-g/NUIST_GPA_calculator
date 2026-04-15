from app.constants import BEND_RELATIONSHIP_1, BEND_RELATIONSHIP_2


def calculate(score):
    # 如果是等级制
    if score in BEND_RELATIONSHIP_2.keys():
        return BEND_RELATIONSHIP_2[score]
    # 如果是分数值
    else:
        for key, value in BEND_RELATIONSHIP_1.items():
            if score >= key:
                return value

