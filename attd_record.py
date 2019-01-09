# coding: UTF-8
import os
import pandas as pd
import numpy as np
import threading


class AttdRecord(threading.Thread):
    def __init__(self, queue, schedule_path, summary_path, save_dir, time_am, time_pm):
        threading.Thread.__init__(self)
        self.queue = queue

        self.schedule_path = schedule_path
        self.summary_path = summary_path
        self.save_dir = save_dir

        self.time_am = pd.to_datetime(time_am)
        self.time_noon = pd.to_datetime('12:00')
        self.time_pm = pd.to_datetime(time_pm)
        self.festival_ratio = 0.7

    def run(self):
        """
        schedule_path: 打卡时间表
        summary_path：月度汇总表
        """
        schedule_path = self.schedule_path
        summary_path = self.summary_path
        save_dir = self.save_dir

        print('schedule_path: %s' % schedule_path)
        print('summary_path: %s' % summary_path)

        # 汇总信息
        summary_xl = pd.ExcelFile(summary_path)
        summary_df = summary_xl.parse(summary_xl.sheet_names[0], header=None)
        summary_columns = summary_df.iloc[3].values
        summary_df = summary_df.drop(labels=[0, 1, 2, 3], axis=0)
        summary_df.columns = summary_columns
        summary_indexes = summary_df.index.values

        # 打卡信息
        schedule_xl = pd.ExcelFile(schedule_path)
        #  sheet_names = schedule_xl.sheet_names
        schedule_df = schedule_xl.parse(
            schedule_xl.sheet_names[0], header=None)

        # eg 打卡时间表 统计日期：2018-11-01 至 2018-11-30
        stats_date_text = ' '.join(schedule_df.iloc[0][0].split()[1:])
        schedule_columns = schedule_df.iloc[2].values
        schedule_df = schedule_df.drop(labels=[0, 1, 2], axis=0)
        schedule_df.columns = schedule_columns
        schedule_indexes = schedule_df.index.values

        names = list()
        user_ids = list()  # user id
        afls = list()  # ask for leave 请假
        evections = list()  # 出差
        absences = list()  # 旷工
        work_days = list()  # 出勤
        forget_counts = list()  # 缺卡（忘记打卡）次数
        late_counts = list()  # 迟到次数
        late_ratios = list()  # 迟到率
        descs = list()  # 描述
        total_latetimes = list()  # 迟到时长

        schedule_values = schedule_df.values
        for summary_index, schedule_index, schedule_value in \
                zip(summary_indexes, schedule_indexes, schedule_values):
            print('summary_index: {}'.format(summary_index))
            print('schedule_index: {}'.format(schedule_index))
            print(schedule_value)

            name = schedule_value[0]
            if name.count('离职') > 0:
                continue

            user_id = schedule_value[4]

            afl = 0
            evection = 0
            absence = 0
            work_day = 0
            forget_count = 0
            late_count = 0
            late_ratio = 0
            desc = ''
            desc_count = 0
            total_latetime = 0

            for schedule_column, item in zip(schedule_columns[5:], schedule_value[5:]):
                print('schedule_column: {}'.format(schedule_column))
                print('item: {}'.format(item))

                if schedule_column in ['六', '日']:
                    continue
                elif ((len(schedule_df) - schedule_df[schedule_column].count()) / len(schedule_df)) >= self.festival_ratio:
                    # 是否是节假日
                    continue
                else:
                    summary = summary_df.loc[summary_index, schedule_column]
                    print('summary: {}'.format(summary))
                    if pd.isna(summary) or pd.isnull(summary):
                        summary = ''

                    if pd.isna(item) or pd.isnull(item) or item.rstrip() == '':  # 没有出勤记录
                        # 判断是出勤，还是请假
                        if schedule_column in summary_columns:
                            if summary.count('旷工') > 0:
                                absence += 1
                            elif summary.count('出差') > 0:
                                evection += 1
                            elif summary.count('假') > 0:
                                afl += 1
                            elif summary.count('外') > 0:
                                evection += 1
                            else:
                                continue
                        else:
                            continue

                    else:  # 有出勤记录，注意只有一次打卡和多次打卡情况
                        work_day += 1

                        times = item.split()
                        day_desc = '%s号（%s%s%s）'
                        am_desc = ''
                        pm_desc = ''

                        is_late_am = False
                        is_late_pm = False
                        if summary.count('上班1迟到') > 0:
                            is_late_am = True

                        if summary.count('上班2迟到') > 0:
                            is_late_pm = True

                        # 上午打卡
                        time_am = None
                        for time in times:
                            time = time[:5]
                            if pd.to_datetime(time) < self.time_noon:
                                time_am = pd.to_datetime(time)  # 上午打卡时间
                                break
                        print('time_am: {}'.format(time_am))

                        if time_am is None:
                            # 检查是否外出等，出差，请假等
                            is_special = False  # 是否为特殊情况
                            if summary.count('出差') > 0:
                                evection += 1
                                is_special = True
                            elif summary.count('假') > 0:
                                afl += 1
                                is_special = True
                            elif summary.count('外') > 0:
                                evection += 1
                                is_special = True

                            if not is_special:
                                forget_count += 1
                                am_desc = '上午缺卡'
                        else:
                            sub_am = time_am - self.time_am
                            if sub_am.total_seconds() > 0:
                                if is_late_am:
                                    late_count += 1
                                    latetime = int(sub_am.total_seconds() / 60)
                                    total_latetime += latetime
                                    am_desc = '上午%s分钟' % str(latetime)

                        # 下午打卡
                        time_pm = None
                        for time in times:
                            time = time[:5]
                            if pd.to_datetime(time) > self.time_noon:
                                time_pm = pd.to_datetime(time)
                                break
                        print('time_pm: {}'.format(time_pm))

                        if time_pm is None:
                            # 检查是否外出等，出差，请假等
                            is_special = False  # 是否为特殊情况
                            if summary.count('出差') > 0:
                                evection += 1
                                is_special = True
                            elif summary.count('假') > 0:
                                afl += 1
                                is_special = True
                            elif summary.count('外') > 0:
                                evection += 1
                                is_special = True

                            if not is_special:
                                forget_count += 1
                                pm_desc += '下午缺卡'
                        else:
                            sub_pm = time_pm - self.time_pm
                            if sub_pm.total_seconds() > 0:
                                if is_late_pm:
                                    late_count += 1
                                    latetime = int(sub_pm.total_seconds() / 60)
                                    total_latetime += latetime
                                    pm_desc = '下午%s分钟' % str(latetime)

                        if am_desc != '' and pm_desc != '':
                            desc_count += 1
                            desc += day_desc % (schedule_column,
                                                am_desc, '、', pm_desc) + ' '
                        elif am_desc != '' or pm_desc != '':
                            desc_count += 1
                            desc += day_desc % (schedule_column,
                                                am_desc, '', pm_desc) + ' '
                        #  if desc_count % 3 == 0:
                            #  desc += '\n'

            if work_day == 0:
                continue

            late_ratio = np.round(late_count / (work_day * 2) * 100, 2)
            late_ratio = str(late_ratio) + '%'

            names.append(name)
            user_ids.append(user_id)
            afls.append(afl)
            evections.append(evection)
            absences.append(absence)
            forget_counts.append(forget_count)
            work_days.append(work_day)
            late_counts.append(late_count)
            total_latetimes.append(total_latetime)
            late_ratios.append(late_ratio)
            descs.append(desc)

        stats_df = pd.DataFrame({
            '姓名': names,
            'UserId': user_ids,
            '请假次数（次）': afls,
            '出差次数（次）': evections,
            '旷工次数（次）': absences,
            '缺卡次数（次）': forget_counts,
            '出勤次数（次）': work_days,
            '迟到次数（次）': late_counts,
            '迟到时长（分钟)': total_latetimes,
            '迟到率': late_ratios,
            '每一次迟到描述': descs
        })

        # 不需要保存索引
        schedule_path_parts = os.path.split(schedule_path)
        schedule_filename_parts = schedule_path_parts[-1].split('_')
        schedule_filename_parts[1] = '考勤汇总'

        save_filename = '_'.join(schedule_filename_parts)
        save_path = os.path.join(save_dir, save_filename)

        stats_df.to_excel(save_path, index=False)
        print('统计完成。')

        self.queue.put('统计完成。')


if __name__ == '__main__':
    time_am = '08:00'
    time_pm = '14:00'
    schedule_path = './data/教务处_打卡时间表_20181101-20181130.xlsx'
    summary_path = './data/教务处_月度汇总_20181101-20181130.xlsx'
    attd_record = AttdRecord(time_am, time_pm)
    attd_record.stats(schedule_path, summary_path)
