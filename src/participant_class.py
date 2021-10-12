import openpyxl
from dataclasses import dataclass


@dataclass
class Participant:
    """Class for participant data object"""
    participant: str
    digit_span_fwd: int
    digit_span_bk: int
    nwr_n: int
    nwr_e: int
    nwr_dc: int
    celf_bin: int
    celf_3pt: int
    lex_tale: int
    poll_sen_rep_bin: int
    poll_sen_rep_3pt: int
    poll_sen_rep_srt: int
    poll_sen_rep_lng: int
    poll_sen_rep_adj: int
    poll_sen_rep_arg: int
    poll_sen_rep_sadj: int
    poll_sen_rep_ladj: int
    poll_sen_rep_sarg: int
    poll_sen_rep_larg: int

    def __init__(self, filepath: str):
        self.fetch_data(filepath)

    def fetch_data(self, filepath: str):
        """
        Fetch data from participant workbook
        :param filepath: str -> absolute file path of source excel file
        """
        workbook = openpyxl.load_workbook(filepath, data_only=True)
        self.participant = workbook['Sum']['b2'].value
        self.digit_span_fwd = workbook['DigitSpan']['e19'].value
        self.digit_span_bk = workbook['DigitSpan']['k19'].value
        self.nwr_n = workbook['NWR_N']['i6'].value
        self.nwr_e = workbook['NWR_L']['i6'].value
        self.nwr_dc = workbook['NWR_D']['i5'].value
        self.celf_bin = workbook['CELF_RS']['c52'].value
        self.celf_3pt = workbook['CELF_RS']['d52'].value
        self.lex_tale = workbook['Sum']['t8'].value
        self.poll_sen_rep_bin = workbook['PollSenRep']['c81'].value
        self.poll_sen_rep_3pt = workbook['PollSenRep']['d81'].value
        self.poll_sen_rep_srt = workbook['PollSenRep']['h4'].value
        self.poll_sen_rep_lng = workbook['PollSenRep']['h5'].value
        self.poll_sen_rep_adj = workbook['PollSenRep']['h6'].value
        self.poll_sen_rep_arg = workbook['PollSenRep']['h7'].value
        self.poll_sen_rep_sadj = workbook['PollSenRep']['h8'].value
        self.poll_sen_rep_ladj = workbook['PollSenRep']['h9'].value
        self.poll_sen_rep_sarg = workbook['PollSenRep']['h10'].value
        self.poll_sen_rep_larg = workbook['PollSenRep']['h11'].value
