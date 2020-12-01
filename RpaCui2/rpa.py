import project 
import mainframe 
import os
import yaml
import pathlib 
import subprocess
import shutil

PY_YAML     = 'yaml'
EXL_APP     = 'Excel.Application'
CSCRIPT     = 'cscript'

RSTCD       = "XXXXXXXXXXXXX"

MOTEKI_CODE='0029005'
S_JIS = 'Shift-JIS'
WRT='w'
RED='r'
NULL_STRING=""
CRLF="\n"

def main(rpa_prj, arg):
    MAC_ENPOINT = 'rpa'

    # 名前の定義 
        # テストの場合、ROO_PRJ_DIR のマスターを使用する

        # ATC_EXE   ... AttacheCase.exeのフルパス
        # SRC_INBOX ... モテキからのメールの受信フォルダ
        # BAK_INBOX ... メール処理後に移動させるフォルダ
        # ORG_MST   ... ユーザが用意したマスターファイル（コピー元）
        # MS_FILE   ... システム入力（処理用）マスターファイル（コピー先）
        # TMP_CSV   ... マスターファイルからモテキ分のみ抽出後

    ATC_EXE = "" 
    SRC_INBOX = "モテキ"                                
    BAK_INBOX = "モテキ保管"

    if hasattr(rpa_prj, 'SYS_PRJ_RPA'):
        myrpa = rpa_prj.SYS_PRJ_RPA
        if "attache_case" in myrpa:
            ATC_EXE = myrpa["attache_case"]
            print("attache_case = " + ATC_EXE)
        if "src_inbox" in myrpa:
            SRC_INBOX = myrpa["src_inbox"]
        if "bak_inbox" in myrpa:
            BAK_INBOX = myrpa["bak_inbox"]
                

    ORG_MST = rpa_prj.USR_PRJ_DIR + "/" + "master.csv"
    #-- TEST ----------------------------------------------------------------#
    if len(arg) > 1:
        if arg[1] == "test":
            ORG_MST = rpa_prj.ROO_PRJ_DIR + "/" + "master.csv"
    #------------------------------------------------------------------------#


    MS_FILE = rpa_prj.USR_PRJ_WRK + "/" + "master.csv"
    RST_FIL = rpa_prj.USR_PRJ_SYS + '/' + 'start_code'
    TMP_CSV = rpa_prj.USR_PRJ_WRK + "/" + "tmp.csv"

    # マスターファイルをworkにコピー
    # テストの場合、確認はしない
    if ORG_MST == rpa_prj.USR_PRJ_DIR + "/" + "master.csv":
        print("確認: 以下のファイルをセットしていますか？ >> " + ORG_MST)
        os.system('PAUSE')

    if not os.path.exists(ORG_MST):
        print("致命的エラー: ファイルがありません >> " + ORG_MST)
        return 0

    shutil.copy(ORG_MST, MS_FILE)








# 添付ファイルの取得
    print("添付ファイルを探しています・・・")
    if not ( len(arg) > 1 and arg[1] == "test" ):
        rtn = subprocess.call(
            CSCRIPT + " " + rpa_prj.USR_PRJ_EXE + "/" + "get_attachment_file.vbs" + " "
                          + SRC_INBOX + " "
                          + rpa_prj.USR_PRJ_WRK + " "
                          + BAK_INBOX 
        )
        if rtn == 1:
            print("警告: 添付ファイルを取得できませんでした")
        if rtn == 2:
            print("警告: 添付ファイルが２ファイル以上ありました")









# 取得した添付ファイルを探す
    # workディレクトリに.atcが１ファイルのみ存在することが条件
    # 作成日時・更新日時等をチェックすることでより精緻化可能
    # テストの場合、ROO_PRJ_DIRの入力ファイルを使用する

    # IN_BASE ... .atcファイルがあった場合、その拡張子を除いたファイル名(hogehoge)
    # IN_EXTO ... .atcファイルがあった場合、その拡張子(.atc)
    # OLG_FIL ... .atcファイルがあった場合、そのファイル名(hogehoge.atc)
    # IN_ATCF ... .atcファイルがあった場合、そのファイルのフルパス

    IN_BASE, IN_EXTO = "", ""
    if not ( len(arg) > 1 and arg[1] == "test" ):
        for f in os.listdir(rpa_prj.USR_PRJ_WRK):
            if ".atc" in f:
                IN_BASE, IN_EXTO = os.path.splitext(f)

        if IN_BASE == "":
            print("致命的エラー: 拡張子.atcを持つファイルがありません")
            print("プログラムを終了します")
            return 0

        OLG_FIL = IN_BASE + IN_EXTO
        IN_ATCF = rpa_prj.USR_PRJ_WRK + "/" + IN_BASE + IN_EXTO
        if OLG_FIL in os.listdir(rpa_prj.USR_PRJ_WRK):
            pass
        else:
            print("致命的エラー: ファイルがありません >> " + IN_ATCF)
            print("プログラムを終了します")
            return 0









    # テストの場合、ROO_PRJ_DIRの入力ファイルを使用する
    # IN_ATCF ... .atcファイルがあった場合、そのファイルのフルパス

    print("添付ファイルを解凍しています・・・")
    if not ( len(arg) > 1 and arg[1] == "test" ):
        if os.path.exists(ATC_EXE):
            pass
        else:
            print("致命的エラー: アッタシェケースがありません >> " + ATC_EXE)
            print("プログラムを終了します")
            return 0
        rtn = subprocess.call(
            ATC_EXE + " "
            + IN_ATCF.replace("/", "\\") + " "
            + "/p=" + rpa_prj.ROO_PRJ_RPA["DECRIPT_PW"] + " "
            + "/de=1"         + " "
            + "/ow=0"         + " "
            + "/opf=0"        + " "
            + "/exit=1"
        )











# 解凍後ファイルを入力ファイル名に変換
    # 解凍後のファイルの拡張子は.xlsになっている
    # テストの場合、ROO_PRJ_DIRの入力ファイルを使用する
    # それをシステムで処理しやすい固定名にリネームする

    # OLG_FI2 ... コピー元（解凍後）ファイル名(hogehoge.xls)
    #             AttacsheCaseは、圧縮ファイル名と同じ名前で解凍する
    #             (同じファイル名で、圧縮するが正しい)
    #             ただし、設定によっては、解凍後ファイル名を固定に出来る可能性
    #             あるので、そのようにした方がいいのかも・・・
    # IN_XLSF ... コピー元（解凍後）ファイルのフルパス
    # SRC_DIR ... コピー元ファイルのあるディレクトリ
    # IN_FILE ... コピー先ファイル（システム入力ファイル）のフルパス

    OLG_FI2 = IN_BASE + ".xls"
    IN_XLSF = rpa_prj.USR_PRJ_WRK + "/" + IN_BASE + ".xls"

    #-- TEST ------------------------------------------------*
    SRC_DIR = rpa_prj.USR_PRJ_WRK
    if ( len(arg) > 1 and arg[1] == "test" ):
        OLG_FI2 = "input.xls" 
        SRC_DIR = rpa_prj.ROO_PRJ_DIR 
        IN_XLSF = rpa_prj.ROO_PRJ_DIR + "/" + OLG_FI2
    #--------------------------------------------------------*

    IN_FILE = rpa_prj.USR_PRJ_WRK + "/" + "input.xls" 

    if OLG_FI2 in os.listdir(SRC_DIR):
        shutil.copy(IN_XLSF, IN_FILE)
    else:
        print("致命的エラー: ファイルがありません >> " + IN_XLSF)
        print("プログラムを終了します")
        return 0










# データ振り分け
#    #   (VBAで処理しやすくするために、各ファイルに振り分けをする)
#    #   (今後、JSONにまとめるよう変更する方向)
#    # 設定ファイルはルートにあるものを適用する
#    # ROO_PRJ_RPA ... ルートにあるYAML OBJECT
#    data = rpa_prj.ROO_PRJ_RPA 
#
#    key = []
#    key.append('gnet1')
#    key.append('gnet2')
#    key.append('gnet3')
#    key.append('siten')
#    key.append('yucho')
#    key.append('taiko')
#    key.append('mitsui')
#    key.append('akita')
#    key.append('joyo')
#    key.append('mizuho')
#    key.append('iraisyo')
#    key.append('asikaga')
#
#    #   キー毎にファイルを作成する(金融機関)
#    for k in key:
#        read_file = rpa_prj.USR_PRJ_WRK + '/' + k
#        with open(read_file, WRT, encoding=S_JIS) as f:
#            sequence = data[k]
#            for scalor in sequence:
#                scalor = scalor.strip()
#                f.write(scalor + '\n')
#
#
#    #   キー毎にファイルを作成する(SIS)
#    key = []
#    key2 = 'sis'
#    key.append('ziei_sinkin')
#    key.append('sinkin')
#    for k in key:
#        read_file = rpa_prj.USR_PRJ_WRK + '/' + k
#        with open(read_file, WRT, encoding=S_JIS) as f:
#            sequence = data[key2][k]
#            for scalor in sequence:
#                scalor = scalor.strip()
#                f.write(scalor + CRLF)
#    read_file = rpa_prj.USR_PRJ_WRK + '/' + key2
#    with open(read_file, WRT, encoding=S_JIS) as f:
#        for k in key:
#            sequence = data[key2][k]
#            for scalor in sequence:
#                scalor = scalor.strip()
#                f.write(scalor + CRLF)
# 
#    key = 'iraisyo'
#    write_file = rpa_prj.USR_PRJ_WRK + '/' + key
#    with open(write_file, WRT, encoding=S_JIS) as f:
#        sequence = data[key]
#        for scalor in sequence:
#            scalor = scalor.strip()
#            f.write(scalor + CRLF)









    # 停止リストからcsvを作成 ------------------------------------------------*
    rtn = subprocess.call(
        CSCRIPT + " " + rpa_prj.MAC_INVOKER + " "
                      + rpa_prj.RPA_XLM     + " "
                      + MAC_ENPOINT + " "
                      + "list_caller"
    )
    #-------------------------------------------------------------------------*










    # tmp.csv の作成 ---------------------------------------------------------*
    in_file = MS_FILE
    ot_file = TMP_CSV
    if os.path.isfile(ot_file):
        os.remove(ot_file)
    with open(ot_file, WRT, encoding=S_JIS) as otf:
        with open(in_file, RED, encoding=S_JIS) as inf:
            header_flg = False
            for line in inf:
                trm_line = mainframe.trim_host_datas(line)
                datas = trm_line.split(',')
                if not header_flg:
                    otf.write(trm_line + CRLF)
                    header_flg = True
                else:
                    if str(datas[2]) == MOTEKI_CODE:
                        otf.write(trm_line + CRLF)
    #-------------------------------------------------------------------------*






    # tmp.csv とteishi_list_tmp.csv のcompare --------------------------------*
    ms_file = TMP_CSV
    in_file = rpa_prj.USR_PRJ_WRK + '/' + 'teishi_list_tmp.csv'
    ot_file = rpa_prj.USR_PRJ_WRK + '/' + 'teishi_result.csv'
    o2_file = rpa_prj.USR_PRJ_WRK + '/' + 'teishi_write.csv'
    err_msg = [''] * 16
    err_msg[4]  = '顧客コードなし'
    err_msg[13] = '口座名義が不一致'
    err_msg[14] = '振替金額が不一致'
    err_msg[15] = 'チェックＯＫ'

    header_flg=True 
    err_flg=4

    rst_flg, rst_cod = get_restart_code(RST_FIL)
    otf = open(ot_file, WRT, encoding=S_JIS)
    o2f = open(o2_file, WRT, encoding=S_JIS)
    msf = open(ms_file, RED, encoding=S_JIS)
    inf = open(in_file, RED, encoding=S_JIS)
    for in_line in inf:
        in_line = in_line.replace(" ", NULL_STRING)
        in_datas = in_line.split(',')

        for mst_line in msf:
            mst_line = mst_line.replace('"', NULL_STRING)
            mst_line = mst_line.replace('=', NULL_STRING)
            mst_line = mst_line.replace(CRLF, NULL_STRING)

            if header_flg:
                otf.write(mst_line + CRLF)
                header_flg=False
                continue

            mst_datas = mst_line.split(',')
            mst_datas[4]  = mst_datas[4].rstrip().lstrip("0").zfill(8)
            mst_datas[13] = mst_datas[13].rstrip() 
            mst_datas[14] = mst_datas[14].rstrip() 

            if in_datas[4] == mst_datas[4]:
                err_flg=13
                if in_datas[13] == mst_datas[13].replace(" ", NULL_STRING):
                    err_flg=14
                    if in_datas[14].replace("\\", NULL_STRING) == mst_datas[14].replace(" ", NULL_STRING):
                        err_flg=15
                break
            else:
                continue

        if err_flg <= 4:
            out_line=in_line.rstrip()
            print("!!! 注意 マスタにデータなし!!!")
            print("   ==> " + in_line.rstrip())
        else:
            out_line = format_line(mst_line)
        out_line+=","+str(err_flg)+","+err_msg[err_flg]

        #--------------------------------------
        if rst_flg:
            otf.write(out_line + CRLF)
        else:
            if mst_datas[4] == rst_cod:
                rst_flg = True
        #--------------------------------------

        if err_flg > 4:
            o2f.write(out_line + CRLF)

        err_flg=4
        msf.seek(0,0)
    otf.close()
    o2f.close()
    msf.close()
    inf.close()
    #-------------------------------------------------------------------------*






    # 依頼書ファイルをwork にコピー ------------------------------------------*
    b, _ = get_restart_code (RST_FIL)
    if b:
        key = "iraisyo"
        data = rpa_prj.ROO_PRJ_RPA
        sequence = data[key]
        for scalor in sequence:
            scalor = scalor.strip()
            if os.path.exists(rpa_prj.USR_PRJ_DIR + "/" + scalor):
                src = rpa_prj.USR_PRJ_DIR + "/" + scalor
                dst = rpa_prj.USR_PRJ_WRK + "/" + scalor
                shutil.copy(src, dst)
            else:
                print("警告: ファイルが存在しません >> " + src)
    #-------------------------------------------------------------------------*







    # 停止依頼書の作成 -------------------------------------------------------*
    rtn = subprocess.call(
        CSCRIPT + " " + rpa_prj.MAC_INVOKER + " "
                      + rpa_prj.RPA_XLM     + " "
                      + MAC_ENPOINT     + " "
                      + "iraisho_caller"
    )
    #-------------------------------------------------------------------------*






    # 送付明細をエクセル化 ---------------------------------------------------*
    if b:
        rtn = subprocess.call(
            CSCRIPT + " " + rpa_prj.MAC_INVOKER + " "
                          + rpa_prj.RPA_XLM     + " "
                          + MAC_ENPOINT     + " "
                          + "save_master_as_xlsx"
        )

    rtn = subprocess.call(
        CSCRIPT + " " + rpa_prj.MAC_INVOKER + " "
                      + rpa_prj.RPA_XLM     + " "
                      + MAC_ENPOINT     + " "
                      + "save_list"
    )
    #-------------------------------------------------------------------------*







    # 最終の場合、リセットファイルに終了コード送る ---------------------------*
    if ( len(arg) > 1 and arg[1] == "end" ):
        with open(RST_FIL, WRT, encoding=S_JIS) as ref:
            ref.write(RSTCD)
    #-------------------------------------------------------------------------*







    # 停止リスト、送付明細、停止依頼書を印刷 ---------------------------------*
    # (最終の場合)
    b, _ = get_restart_code (RST_FIL)
    if b:
        rtn = subprocess.call(
            CSCRIPT + " " + rpa_prj.MAC_INVOKER + " "
                          + rpa_prj.RPA_XLM     + " "
                          + MAC_ENPOINT     + " "
                          + "sort_teishi_list"
        )
        rtn = subprocess.call(
            CSCRIPT + " " + rpa_prj.MAC_INVOKER + " "
                          + rpa_prj.RPA_XLM         + " "
                          + MAC_ENPOINT         + " "
                          + "print_teishi_list"
        )
        rtn = subprocess.call(
            CSCRIPT + " " + rpa_prj.MAC_INVOKER + " "
                          + rpa_prj.RPA_XLM     + " "
                          + MAC_ENPOINT     + " "
                          + "print_soufu_meisai"
        )
        rtn = subprocess.call(
            CSCRIPT + " " + rpa_prj.MAC_INVOKER + " "
                          + rpa_prj.RPA_XLM     + " "
                          + MAC_ENPOINT     + " "
                          + "print_summary"
        )
        rtn = subprocess.call(
            CSCRIPT + " " + rpa_prj.MAC_INVOKER + " "
                          + rpa_prj.RPA_XLM     + " "
                          + MAC_ENPOINT     + " "
                          + "print_sheet"
        )
    #-------------------------------------------------------------------------*






    # 停止依頼書等をバックアップにコピー -------------------------------------*
    key = "iraisyo"
    data = rpa_prj.ROO_PRJ_RPA 
    sequence = data[key]
    for scalor in sequence:
        scalor = scalor.strip()
        org = rpa_prj.USR_PRJ_DIR + "/" + scalor
        src = rpa_prj.USR_PRJ_WRK + "/" + scalor
        dst = rpa_prj.USR_PRJ_BAK + "/" + scalor
        if os.path.exists(org):
            shutil.copy(src, dst)
        else:
            print("警告: ファイルが存在しません >> " + org)
            print("　　　このため、ファイルを作成しません >> " + dst)
    
    shutil.copy(rpa_prj.USR_PRJ_WRK + "/" + "modified_master.xlsx", rpa_prj.USR_PRJ_BAK + "/" + "加工済送付明細.xlsx")
    shutil.copy(rpa_prj.USR_PRJ_WRK + "/" + "summary.csv", rpa_prj.USR_PRJ_BAK + "/" + "件数集計.csv")
    #-------------------------------------------------------------------------*
    






    # ファイルの削除 ---------------------------------------------------------*
    if os.path.exists(rpa_prj.USR_PRJ_WRK + "/" + "modified_teishi_result.xlsx"):
        os.remove(rpa_prj.USR_PRJ_WRK + "/" + "modified_teishi_result.xlsx")

    if b:
        for f in os.listdir(rpa_prj.USR_PRJ_WRK):
            #print(rpa_prj.USR_PRJ_WRK + "/" + f)
            os.remove(rpa_prj.USR_PRJ_WRK + "/" + f)
    #-------------------------------------------------------------------------*




    return 0

        








def get_restart_code (RST_FIL):
    with open(RST_FIL, RED, encoding=S_JIS) as ref:
        code=ref.read().rstrip()
        if code == RSTCD:
            flag=True
            code=RSTCD
        else:
            flag=False
            code=code.lstrip("0").zfill(8)
        return flag, code

def format_line(in_line):
    out_line = ''
    datas = in_line.split(',')
    for idx, data in enumerate(datas):
        if idx == (6 | 15):
            data = data.rstrip()
        else:
            data = data.strip()
        out_line += data + ','
    out_line = out_line.rstrip(',')
    return out_line

if __name__ == '__main__':
    main()
