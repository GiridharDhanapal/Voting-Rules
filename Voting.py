from openpyxl import load_workbook

def tieBreaker(preferences, score, tieBreak):
    try:
        alternates = []
        for keys in score.keys():
            if score[keys] == max(score.values()):
                alternates.append(keys)
        if list(score.values()).count(max(score.values()))>1:
            if isinstance(tieBreak, int):
                if tieBreak not in preferences.keys():
                    raise ValueError(f'Invalid agent number')
                else:
                    for i in preferences[tieBreak]:
                        if i in alternates:
                            return i
            elif tieBreak == "max":
                return max(alternates)
            elif tieBreak == "min":
                return min(alternates)
            else:
                raise ValueError(f'Invalid TieBreak')
        else:
            return alternates[0]
    except:
        raise ValueError(f"Incorrect Tie Break value")
        
def generatePreferences(values):
    preferences = {}
    agent = 0
    for row in values.rows:
        preference = [data.value for data in row]
        copy_preference = preference[:]
        copy_preference.sort()
        new_list = []
        for i in copy_preference[::-1]:
            new_list.append(len(preference)-preference[::-1].index(i))
            preference[len(preference)-preference[::-1].index(i)-1] = -1
        agent += 1
        preferences.update({agent : new_list})
    return preferences

def dictatorship(preferenceProfile, agent):
    try:
        if agent not in preferenceProfile.keys():
            raise ValueError(f'Invalid agent number')
        else:
            return preferenceProfile[agent][0]
    except:
        raise ValueError(f'Invalid agent number')
    
def scoringRule(preferences, scoreVector, tieBreak):
    try:
        if len(scoreVector) != len(preferences[1]):
            raise ValueError(f'Invalid scoreVector length number')
        else:
            scoreVector.sort()
            scoreVector = scoreVector[::-1]
            score = {}
            for preference in preferences.values():
                for i in range(len(preference)):
                    if preference[i] in score.keys():
                        score.update({preference[i]:scoreVector[i]+score[preference[i]]})
                    else:
                        score.update({preference[i]:scoreVector[i]})
            return tieBreaker(preferences, score, tieBreak)
    except: 
        print("Incorrect input")
        return False
        
def plurality(preferences, tieBreak):
    scoreVector = [1]+[0 for i in range(len(list(preferences.values())[0])-1)]
    scoreVector.sort()
    scoreVector = scoreVector[::-1]
    score = {}
    for preference in preferences.values():
        for i in range(len(preference)):
            if preference[i] in score.keys():
                score.update({preference[i]:scoreVector[i]+score[preference[i]]})
            else:
                score.update({preference[i]:scoreVector[i]})
    return tieBreaker(preferences, score, tieBreak)
  
def veto(preferences, tieBreak):
    scoreVector = [1 for i in range(len(list(preferences.values())[0])-1)]+[0]
    scoreVector.sort()
    scoreVector = scoreVector[::-1]
    point = {}
    for preference in preferences.values():
        for i in range(len(preference)):
            if preference[i] in point.keys():
                point.update({preference[i]:scoreVector[i]+point[preference[i]]})
            else:
                point.update({preference[i]:scoreVector[i]})
    return tieBreaker(preferences, point, tieBreak)
    
def borda(preferences, tieBreak):
    scoreVector = list(range(len(list(preferences.values())[0])-1,-1,-1))
    scoreVector.sort()
    scoreVector = scoreVector[::-1]
    score = {}
    for preference in preferences.values():
        for i in range(len(preference)):
            if preference[i] in score.keys():
                score.update({preference[i]:scoreVector[i]+score[preference[i]]})
            else:
                score.update({preference[i]:scoreVector[i]})
    return tieBreaker(preferences, score, tieBreak)

def harmonic(preferences, tieBreak):
    scoreVector = [1/i for i in range(1,len(list(preferences.values())[0])+1)]
    scoreVector.sort()
    scoreVector = scoreVector[::-1]
    score = {}
    for preference in preferences.values():
        for i in range(len(preference)):
            if preference[i] in score.keys():
                score.update({preference[i]:scoreVector[i]+score[preference[i]]})
            else:
                score.update({preference[i]:scoreVector[i]})
    return tieBreaker(preferences, score, tieBreak)

def STV(preferences, tieBreak):
    scoreVector = [1]+[0 for i in range(len(list(preferences.values())[0])-1)]
    scoreVector.sort()
    scoreVector = scoreVector[::-1]
    score = {}
    for preference in preferences.values():
        for i in range(len(preference)):
            if preference[i] in score.keys():
                score.update({preference[i]:scoreVector[i]+score[preference[i]]})
            else:
                score.update({preference[i]:scoreVector[i]})
    return tieBreaker(preferences, score, tieBreak)

def rangeVoting(values, tieBreak):
    columns = []
    sums = {}
    for row in values.rows:
        valuation = [data.value for data in row]
        columns.append(valuation)
    for i in range(len(columns[0])):
        sum = 0
        for j in range(len(columns)):
            sum += columns[j][i]
        sums.update({i+1:sum})
    return tieBreaker(generatePreferences(values), sums, tieBreak)
     