"""
!/usr/bin/env python3
# -*- coding: utf-8 -*-
Author: Kai Lu, Sun Yat-sen University, Dec. 25, 2020
Copyright reserved
reference:
https://www.jianshu.com/p/dfebe1f4ddcf
"""
from email.parser import Parser
from email.header import decode_header
import poplib
import random
import warnings
from email.header import Header
from email.mime.text import MIMEText
from email.utils import parseaddr, formataddr
import smtplib
import pandas as pd
import os


def _format_addr(s):
    name, addr = parseaddr(s)
    return formataddr((Header(name, 'utf-8').encode(), addr))


def send_email(to_name, to_addr, send_msg):
    email_info_filename = 'email_inf.csv'
    email_info = pd.read_csv(email_info_filename, sep=',')
    email = email_info.iloc[0, 0]
    authentication_code = email_info.iloc[0, 1]

    from_addr = email
    smtp_server = 'smtp.163.com'
    msg = MIMEText(send_msg, 'plain', 'utf-8')
    msg['From'] = _format_addr('EIT5641-教师<%s>' % from_addr)
    msg['To'] = _format_addr(f'EIT5641-学生-{to_name}<{to_addr}>')
    msg['Subject'] = Header('你对本组同学的打分不成功，下次请用确认过的邮箱、按要求格式发送，谢谢！', 'utf-8').encode()

    server = smtplib.SMTP(smtp_server, 25)
    server.set_debuglevel(0)
    server.login(from_addr, authentication_code)
    server.sendmail(from_addr, [to_addr], msg.as_string())
    server.quit()


def decode_msg_header(header):
    """
    解码头文件
    :param header: 需解码的内容
    :return:
    """
    value, charset = decode_header(header)[0]
    if charset:
        value = value.decode(charset)
    return value


class Fetch_Server:
    def __init__(self):
        # 输入POP3服务器地址:
        pop3_server = 'pop.163.com'
        # 连接到POP3服务器:
        self.active_server = poplib.POP3(pop3_server)
        # 可以打开或关闭调试信息:
        self.active_server.set_debuglevel(0)
        # 输入邮件地址, 口令:
        email_info_filename = 'email_inf.csv'
        # email_info = pd.read_csv(email_info_filename, sep=',')
        # email = email_info.iloc[0, 0]
        # authentication_code = email_info.iloc[0, 1]
        # 身份认证:
        self.active_server.user(pd.read_csv(email_info_filename, sep=',').iloc[0, 0])
        self.active_server.pass_(pd.read_csv(email_info_filename, sep=',').iloc[0, 1])

    def quit(self):
        self.active_server.quit()


def prepare_rate_grp():
    # read class table and group table
    _class_table = pd.read_excel('class_list_2020.xls', convert_float=False)
    _group_table = pd.read_excel('presentation_list.xls', convert_float=False)
    # extract group number and email addresses
    _group_num_list = _group_table.组号.tolist()
    _qualified_email_adr_list = _class_table.邮箱.tolist()
    # add email addresses of instructor, the 2nd and 3rd ones for test
    _qualified_email_adr_list.append('lukai86@mail.sysu.edu.cn')
    _qualified_email_adr_list.append('kai.lu@my.cityu.edu.hk')
    _group_table.set_index('组号', inplace=True)
    _group_table.sort_values(by='是否展示', inplace=True)
    # class_table.set_index('序号', inplace=True)
    return _group_num_list, _qualified_email_adr_list, _group_table


def choose_rate(to_present_grp_table, _rating_results_folder):
    all_csv_file_name = 'All_Groups' + '.csv'
    all_csv_file_dir = os.path.join(_rating_results_folder, all_csv_file_name)
    rating_table = pd.DataFrame()
    to_present_list = to_present_grp_table.index.tolist()
    finished_count = 0
    to_present_grp_count = len(to_present_list)
    while to_present_grp_count != 0:
        num_key = input("按任意键选下一组(按'q'退出程序):")
        if num_key == "q":
            warnings.warn("选组过程中断")
            break
        else:
            to_present_grp_index = random.randint(0, to_present_grp_count - 1)
            next_grp_index = int(to_present_list[to_present_grp_index])
            csv_file_name = 'Group' + str(next_grp_index) + '.csv'
            csv_file_dir = os.path.join(_rating_results_folder, csv_file_name)
            if os.path.exists(csv_file_dir):
                finished_count += 1
                to_present_grp_count -= 1
                to_present_list.pop(to_present_grp_index)
                print(f'第{next_grp_index}组已完成, 下一组')
            else:
                print(f'第{next_grp_index}组请上台作报告 \n',
                      to_present_grp_table.loc[next_grp_index, ['组员1', '组员2', '选题', '是否展示']])
                print(f"本组演示结束后，请同学们发送自己的评分到\n"
                      f"lecture_sysu@163.com，\n "
                      f"标题格式为: \n"
                      f"{next_grp_index}-分数\n"
                      f"分数取整，0=<分数<=100，不用写邮件正文")
                score_key = input("按任意键显示本组成绩(按'q'退出程序):")
                if score_key == "q":
                    warnings.warn("分数统计过程中断")
                    break
                else:
                    rating_table[next_grp_index] = collect_rate(next_grp_index, csv_file_dir)
                    rating_table.to_csv(all_csv_file_dir)
                    finished_count += 1
                    to_present_grp_count -= 1
                    to_present_list.pop(to_present_grp_index)
    else:
        print(f"本类别所有组的报告均已完成, 共有{finished_count}组同学做了报告.")

    rating_table.to_csv(all_csv_file_dir)
    return rating_table


def collect_rate(current_grp_index, csv_file_dir):
    _rating_list = pd.DataFrame()
    mail_server_class = Fetch_Server()
    mail_server = mail_server_class.active_server
    # stat()返回邮件数量和占用空间:
    msg_num = mail_server.stat()[0]
    print(f'本组同学共收到{msg_num}封打分email')
    if int(msg_num) > 0:
        for mail_index in [msg_num + 1 - item for item in range(1, msg_num + 1)]:
            # 获取邮件, 注意索引号从1开始
            # lines存储了邮件的原始文本的每一行,
            # 可以获得整个邮件的原始文本:
            lines = mail_server.retr(mail_index)[1]
            msg_content = b'\r\n'.join(lines).decode('utf-8')
            # 解析出邮件:
            msg = Parser().parsestr(msg_content)
            # extract sender email
            sender_content = msg["From"]
            # parseaddr()函数返回的是一个元组(real_name, emailAddress)
            sender_real_name, sender_adr = parseaddr(sender_content)
            # # 将加密的名称进行解码
            # sender_real_name = decode_msg_header(sender_real_name)
            if sender_adr in qualified_email_adr_list:
                # extract subject
                msg_header = msg["Subject"]
                # 对头文件进行解码
                msg_header = decode_msg_header(msg_header)
                if '-' in msg_header:
                    grp_rate_info = msg_header.split('-')
                    if len(grp_rate_info) == 2:
                        try:
                            grp_index = int(grp_rate_info[0])
                            if int(grp_index - current_grp_index) == 0:
                                try:
                                    rate = int(grp_rate_info[1])
                                    if 0 <= rate <= 100:
                                        # print(grp_index, rate)
                                        _rating_list.loc[sender_adr, grp_index] = rate
                                    else:
                                        print('发现一张废票, 分值必须在[0, 100]区间')
                                        send_msg = "请确保打分范围是否在0和100之间"
                                        send_email(sender_real_name, sender_adr, send_msg)
                                except Exception as e:
                                    print(e)
                                    print('发现一张废票, 分值必须是数字')
                                    send_msg = "打分必须是数字"
                                    send_email(sender_real_name, sender_adr, send_msg)
                            else:
                                print('发现一张废票, 打分对象只可以是本组')
                                send_msg = "只可为当前组打分"
                                send_email(sender_real_name, sender_adr, send_msg)
                        except Exception as e:
                            print(e)
                            print('发现一张废票, 组号只可以是数字')
                            send_msg = "组号必须是数字"
                            send_email(sender_real_name, sender_adr, send_msg)
                    else:
                        print('发现一张废票, email标题必须是组号-分数的形式')
                        send_msg = "请严格遵循打分格式'组号-分数'"
                        send_email(sender_real_name, sender_adr, send_msg)
                else:
                    print('发现一张废票, 标题中必须包含-')
                    send_msg = "请确保包含了符号'-'"
                    send_email(sender_real_name, sender_adr, send_msg)
            else:
                print('发现一张废票, 必须使用确认过的email发送打分结果')
                send_msg = "请严格使用中大官邮或者授课教师确认过的替代邮箱发信"
                send_email(sender_real_name, sender_adr, send_msg)
            # 可以根据邮件索引号直接从服务器删除邮件:
            mail_server.dele(mail_index)
        _rating_list.to_csv(csv_file_dir)
        print(f'第{current_grp_index}组同学得到的平均分{_rating_list[current_grp_index].mean()}, '
              f'标准差为{_rating_list[current_grp_index].std()}')
    else:
        warnings.warn('没有收到对本组同学的打分')
    # 关闭连接:
    mail_server.quit()
    # print(_rating_list)
    return _rating_list[current_grp_index]


if __name__ == '__main__':
    # prepare the index and columns of the rate table, import the group table
    group_num_list, qualified_email_adr_list, group_table = prepare_rate_grp()
    registered_grp_table = group_table[group_table['是否展示'] == '是']
    non_registered_grp_table = group_table[group_table['是否展示'] != '是']

    rating_results_folder = 'Rating_Results'
    if not os.path.exists(rating_results_folder):
        os.makedirs(rating_results_folder)

    # start with registered groups
    rating_table1 = choose_rate(registered_grp_table, rating_results_folder)
    # continue with non_registered groups
    rating_table2 = choose_rate(non_registered_grp_table, rating_results_folder)
