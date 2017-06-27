#! python3
# -*- coding: utf-8 -*-
"""Index: A collection of row indices for the template excel file"""
class Index(object):

    def __init__(self):
        self.IB_Query = { 'Canada': 17, 'Bahamas': 20, 'Turks_and_Caicos': 23, 'Cayman': 27,
                          'Costa_Rica': 30, 'Jamaica': 33, 'Belize': 36, 'Guyana': 40, 'Antigua': 43,
                          'Anguilla': 46, 'St_Kitts': 49, 'Dominica': 52, 'St_Lucia': 55, 'St_Vincent': 58,
                          'Grenada': 61, 'St_Maarten': 64, 'Barbados': 67, 'BVI': 70,
                          'Trinidad': 73, 'Dominican_Republic': 76, 'Puerto_Rico': 80, 'USVI': 83 }

        self.Query = { 'Canada': 18, 'Bahamas': 21, 'Turks_and_Caicos': 24, 'Cayman': 28,
                       'Costa_Rica': 31, 'Jamaica': 34, 'Belize': 37, 'Guyana': 41, 'Antigua': 44,
                       'Anguilla': 47, 'St_Kitts': 50, 'Dominica': 53, 'St_Lucia': 56, 'St_Vincent': 59,
                       'Grenada': 62, 'St_Maarten': 65, 'Barbados': 68, 'BVI': 71, 'Trinidad': 74,
                       'Dominican_Republic': 77, 'Puerto_Rico': 81, 'USVI': 84 }

        self.Debt_Mgr = { 'Bahamas': 92, 'Costa_Rica': 101, 'Jamaica': 93, 'Dominican_Republic': 97,
                          'Trinidad': 96, 'Barbados': 95, 'St_Lucia': 94, 'Mexico': 102, 'Chile': 100,
                          'USVI': 99, 'Puerto_Rico': 98 }

        self.HORIZON_BANK01_ACQUIRE = { 'Puerto_Rico': 88 }
        self.HORIZON_BANK01_BOSS    = { 'Puerto_Rico': 87 }
        self.HORIZON_BANK15_BOSS    = { 'Puerto_Rico': 89 }

        self.Company_1 = { 'Bahamas_Trust': 105, 'Cayman': 112 }
        self.Company_2 = { 'Bahamas_Trust': 106, 'Jamaica': 113 }
        self.Company_3 = { 'Bahamas_Trust': 107 }
        self.Company_6 = { 'Bahamas_Trust': 108 }
        self.Company_7 = { 'Bahamas_Trust': 109 }
        self.Company_8 = { 'Bahamas_Trust': 110 }
        self.Company_9 = { 'Bahamas_Trust': 111 }

        self.A = { 'Puerto_Rico': 79 }
        self.B = { 'Belize': 35, 'USVI': 82 }
        self.C = { 'Bahamas': 19, 'Dominica': 51 }
        self.D = { 'Cayman': 26, 'Dominican_Republic': 75 }
        self.E = { 'Turks_and_Caicos': 22 }
        self.F = { 'St_Maarten': 63 }
        self.G = { 'St_Kitts': 48 }
        self.H = { 'Trinidad': 72 }
        self.K = { 'Barbados': 66 }
        self.L = { 'Guyana': 39 }
        self.M = { 'St_Vincent': 57 }
        self.N = { 'Grenada': 60 }
        self.O = { 'Canada': 16 }
        self.P = { 'St_Lucia': 54 }
        self.R = { 'Costa_Rica': 29 }
        self.S = { 'Antigua': 42 }
        self.U = { 'Anguilla': 45 }
        self.V = { 'BVI': 69 }

        self.DBA = { 'Jamaica': 32 }
