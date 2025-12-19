@echo off
setlocal enabledelayedexpansion

:PASSWORD_INPUT
cls
echo ���s�m�F

set /p password="�p�X���[�h����͂��Ă�������: "

cd /d "%~dp0program"

rem �p�X���[�h����
python new_automation.py !password! --verify
if errorlevel 1 (
    echo �p�X���[�h���Ⴂ�܂��B�ēx���͂��Ă��������B
    timeout /t 2 > nul
    goto PASSWORD_INPUT
)

rem �p�X���[�h���������ꍇ�̂ݎ��s�m�F��\��
echo.
echo ============================================
echo �y��������z
echo - Escape �L�[�������Ə����������I�����܂�
echo - ���̏����Ń_�E�����[�h����уX�N���[���V���b�g���ꂽ�t�@�C���́A�S��"\Downloads\pdf"�ɕۑ�����܂��B
echo - �K��N�����̉���҂Ɋւ��ẮA��قǕʒS���҂��Ή����܂��̂ŁA�u���Ή��v�t���O�̂܂܉������Ȃ��đ��v�ł��B
echo.
echo �y�m�F�����z
echo - �u�_�E�����[�h�v�t�H���_�Ɂupef�v�t�H���_�͂���܂����H
echo - Edge�̐ݒ�ŁA�_�E�����[�h��̃t�H���_����L�́updf�v�ɕύX���܂������H
echo ���ڂ����͓��t�H���_�ւɂ������������������������B
echo.
echo �y���ӎ����z
echo �}�E�X����ʂ̎l���ɒu���Ă����ԂŎ��s�����
echo �G���[�ɂȂ�܂��B�l�����痣���čēx���s���Ă��������B
echo.
echo �H�Ƀ��W������Y���Ȃ�����҂����܂����A
echo ���̏ꍇ�͑���ɉ���ҏ���ʂ������ŃX�N�V������܂��B
echo ============================================
echo.
choice /c YN /m "�v���O���������s���܂����H(Y=�͂� / N=������)"
if errorlevel 2 (
    echo �v���O�������I�����܂�
    pause
    exit
)

rem ���s�m�F��Yes�̏ꍇ�̂ݎ��s
cls
python new_automation.py !password!
pause