Attribute VB_Name = "Categories"
Option Explicit

' �D�������L�^�̊e���ږ��̃��X�g��Ԃ�
Function InspectRecCategories()

    Dim Categories As Object
    Set Categories = CreateObject("scripting.Dictionary")
    
    ' �����J�e�S���[�̃��X�g��A�z�z��ɐݒ肷��
    ' ��) "stat"= ���ʗp������, "��"= �D�������L�^��ł̌���������
    ' ���ӁF "���s" �� "���s" �͓��̓~�X���N���₷��
    Categories.Add "stat", "��"
    Categories.Add "year", "�N�x"
    Categories.Add "refNum", "��"
    Categories.Add "shipName", "�D��"
    Categories.Add "shipType", "�D�����"
    Categories.Add "owner", "���L��"
    Categories.Add "captainName", "�D����"
    Categories.Add "delegater", "�Ϗ���"
    Categories.Add "delegateStaff", "�S��"
    Categories.Add "inspectType", "�������"
    Categories.Add "clause", "��"
    Categories.Add "receiptDate", "��t��"
    Categories.Add "repNo", "�Ӓ菑�ԍ�"
    Categories.Add "repNoCreateDate", "���s��"
    Categories.Add "concurrentInspection", "���s����"
    Categories.Add "shipyard", "���D��"
    Categories.Add "inspectDate", "�����"
    Categories.Add "unDocking", "���ˁE�o����"
    Categories.Add "prevUndocking", "�O�񉺉ˁE�o����"
    Categories.Add "prevInspection", "�O�񌟍�"
    Categories.Add "prevRepNo", "�O��Ӓ菑"
    Categories.Add "grossT", "���g����"
    Categories.Add "length", "�o�^��"
    Categories.Add "breadth", "�S��"
    Categories.Add "depth", "�S�[�^�^�["
    Categories.Add "shaftDia", "���a"
    Categories.Add "propellerNum", "����"
    Categories.Add "propellerMaterial", "�ގ�"
    Categories.Add "propellerDia", "�_�C��"
    Categories.Add "propellerPitch", "�s�b�`"
    Categories.Add "accidentDetail", "���̓��e"
    Categories.Add "marineAccidentReport", "�C��񍐏�"
    Categories.Add "repairAmount", "�C�U�z"
    Categories.Add "location", "���_"
    Categories.Add "kmsStaff", "�S����"
    Categories.Add "inspectDateOther", "�{����(����ȍ~)"
    Categories.Add "remark", "���L����"
    
    Set InspectRecCategories = Categories

End Function

