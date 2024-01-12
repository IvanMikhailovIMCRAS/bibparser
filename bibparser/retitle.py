class SmartStr(str):
    def capitalize(cls):
        if cls:
            return cls[0].upper() + cls[1:]
        else:
            return ''
        
    def lower(cls):
        if cls:
            if len(cls) == 1:
                return cls[0].lower()
            else:
                if cls[1] == cls[1].upper():
                    return cls
                else:
                   return cls[0].lower() + cls[1:] 
        else:
            return ''


def retitle(title: str) -> str:
    lst = list(map(SmartStr, title.split()))
    lst = [word.lower() for word in lst]
    lst[0] = lst[0].capitalize()
    lst = [lst[0]] + [lst[i].capitalize() if lst[i-1][-1]=='.' else lst[i] for i in range(1,len(lst))]
    return str(' '.join(lst))

if __name__ == '__main__':
    print(retitle("a NMR Study of A-type Polyacid-N-propyl by MD Simulation"))
    
