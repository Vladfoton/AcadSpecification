class Parameters:
    def __init__(self, acad_block:object ):

        self.constr_name = constr_name  # Диафрагма Д1
        self.arm_standart : dict =  self.set_arm_standart(arm_standart)  # {"":('ГОСТ 34028-2016', 'А500C'),"г":('ГОСТ 34028-2016', 'А240'),"в":('ГОСТ 6727-80', 'ВрI),



    @staticmethod
    def set_arm_standart(arm_standart:str)->dict:
        return {arm_data[0]: (arm_data[1], arm_data[2]) for arm_data in arm_standart.split(',')}

'''
        klass0: A500C
        projectcode: 01 - 22 - 14 - 1 / 1 - КЖО
        multipleVRS: 1.1
        workingfolder: Д1
        Detailsfile: VDд.xlsx
        SpecificationType: 0
        VRS_Type: 1
        SpecificationHead: 1
'''
        pass

    def __str__(self):
        pass

    def add(self, textstring: str):
        pass
