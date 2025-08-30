import random
import pandas as pd
from typing import List, Dict, Tuple

class Student:
    def __init__(self, student_id: str, name: str, gender: str, class_name: str):
        self.student_id = student_id
        self.name = name
        self.gender = gender
        self.class_name = class_name
        self.group = None
        self.seat_number = None
        
    def __str__(self):
        return f"{self.name}({self.gender}, {self.class_name}) - 组{self.group} 座位{self.seat_number}"

class StudentGroupingSystem:
    def __init__(self, excel_file=None):
        # 默认的Excel文件路径
        default_excel_path = '/Users/xushida/Nutstore Files/我的坚果云/大连交通大学(新)/教学/2025秋-最优化导论/教学班1(34人)——物流241、242/第一教学班-选课名单.xlsx'
        self.excel_file = excel_file or default_excel_path
        
        # 读取学生数据
        try:
            if self.excel_file:
                self.students = self._read_students_from_excel(self.excel_file)
            else:
                self.students = self._create_sample_students()
        except Exception as e:
            print(f"读取数据时出错: {e}")
            self.students = self._create_sample_students()
        
        # 获取实际学生总数
        self.total_students = len(self.students)
        
        # 根据实际学生总数调整分组配置
        if self.total_students == 34:
            # 对于35人：A组7人，B/C/E组各6人，D/F组各5人
            self.group_config = {
                'A': 6,
                'B': 6,
                'C': 6,
                'D': 5,
                'E': 6,
                'F': 5
            }
        else:
            # 默认配置：A/B/C/E组各6人，D/F组各5人，总共34人
            self.group_config = {
                'A': 6,
                'B': 6,
                'C': 6,
                'D': 5,
                'E': 6,
                'F': 5
            }
        
        # 验证分组配置总人数
        self.config_total = sum(self.group_config.values())
        if self.config_total != self.total_students:
            print(f"警告：分组配置总人数({self.config_total})与学生总数({self.total_students})不匹配！")
        
    def _read_students_from_excel(self, excel_file: str) -> List[Student]:
        """从Excel文件读取学生数据"""
        try:
            # 使用pandas读取Excel文件
            df = pd.read_excel(excel_file)
            
            # 打印DataFrame的前几行和列名，帮助调试
            print("Excel文件读取成功，数据预览：")
            print(f"列名: {list(df.columns)}")
            print(df.head())
            
            students = []
            
            # 尝试不同的列名组合来匹配学生信息
            # 定义可能的列名组合
            possible_column_mappings = [
                # (学号列, 姓名列, 性别列, 班级列)
                ('学号', '姓名', '性别', '班级'),
                ('student_id', 'name', 'gender', 'class'),
                ('id', 'name', 'sex', 'class'),
                ('序号', '姓名', '性别', '班级'),
            ]
            
            # 尝试找到匹配的列名组合
            selected_mapping = None
            for mapping in possible_column_mappings:
                if all(col in df.columns for col in mapping):
                    selected_mapping = mapping
                    break
            
            if selected_mapping:
                id_col, name_col, gender_col, class_col = selected_mapping
                print(f"使用列名映射: 学号='{id_col}', 姓名='{name_col}', 性别='{gender_col}', 班级='{class_col}'")
                
                # 遍历DataFrame创建Student对象
                for _, row in df.iterrows():
                    # 跳过标题行
                    if str(row[id_col]).strip() in ['学号', '学   号']:
                        continue
                    
                    student_id = str(row[id_col])
                    name = str(row[name_col])
                    gender = str(row[gender_col])
                    class_name = str(row[class_col])
                    
                    # 清理数据（移除可能的空格、换行符等）
                    student_id = student_id.strip()
                    name = name.strip()
                    gender = gender.strip()
                    class_name = class_name.strip()
                    
                    # 只添加有效的学生数据
                    if name and name not in ['姓 名', '姓名'] and gender and gender not in ['性别', '性 别']:
                        students.append(Student(student_id, name, gender, class_name))
            else:
                # 如果没有找到匹配的列名组合，尝试使用位置索引
                print("未找到标准列名，尝试使用位置索引...")
                for idx, row in df.iterrows():
                    # 跳过标题行
                    if idx == 0:
                        continue
                    
                    # 假设前四列分别是学号、姓名、性别、班级
                    if len(row) >= 4:
                        student_id = str(row.iloc[0]).strip()
                        name = str(row.iloc[1]).strip()
                        gender = str(row.iloc[2]).strip()
                        class_name = str(row.iloc[3]).strip()
                        
                        # 只添加有效的学生数据
                        if name and name not in ['姓 名', '姓名'] and gender and gender not in ['性别', '性 别']:
                            students.append(Student(student_id, name, gender, class_name))
            
            print(f"成功读取{len(students)}名学生数据")
            return students
        except Exception as e:
            print(f"读取Excel文件时出错: {str(e)}")
            # 如果读取失败，返回模拟数据
            print("返回模拟数据...")
            return self._create_sample_students()
            
    def _create_sample_students(self) -> List[Student]:
        """创建34名学生的模拟数据"""
        # 物流241班级学生（24人）
        logistics_241 = [
            ("2407070101", "张三", "男", "物流241"),
            ("2407070102", "李四", "男", "物流241"),
            ("2407070103", "王五", "男", "物流241"),
            ("2407070104", "赵六", "男", "物流241"),
            ("2407070107", "钱七", "男", "物流241"),
            ("2407070108", "孙八", "男", "物流241"),
            ("2407070109", "周九", "男", "物流241"),
            ("2407070111", "吴十", "男", "物流241"),
            ("2407070112", "郑十一", "女", "物流241"),
            ("2407070113", "王十二", "女", "物流241"),
            ("2407070114", "李十三", "女", "物流241"),
            ("2407070115", "赵十四", "女", "物流241"),
            ("2407070116", "张十五", "女", "物流241"),
            ("2407070118", "陈十六", "女", "物流241"),
            ("2407070119", "杨十七", "女", "物流241"),
            ("2407070120", "黄十八", "女", "物流241"),
            ("2407070121", "周十九", "女", "物流241"),
            ("2407070122", "吴二十", "女", "物流241"),
            ("2407070123", "郑二一", "男", "物流241"),
            ("2407070124", "王二二", "男", "物流241"),
            ("2407070125", "李二三", "男", "物流241"),
            ("2407070126", "赵二四", "女", "物流241"),
            ("2407070127", "张二五", "男", "物流241"),
            ("2407070128", "陈二六", "女", "物流241"),
        ]
        
        # 物流242班级学生（10人）
        logistics_242 = [
            ("2407070203", "杨二七", "女", "物流242"),
            ("2407070207", "黄二八", "女", "物流242"),
            ("2407070210", "周二九", "女", "物流242"),
            ("2407070215", "吴三十", "男", "物流242"),
            ("2407070217", "郑三一", "男", "物流242"),
            ("2407070221", "王三二", "男", "物流242"),
            ("2407070222", "李三三", "女", "物流242"),
            ("2407070224", "赵三四", "女", "物流242"),
            ("2407070226", "张三五", "男", "物流242"),
            ("2407070227", "陈三六", "男", "物流242"),
        ]
        
        # 合并所有学生
        all_students = logistics_241 + logistics_242
        
        # 创建Student对象列表
        return [Student(student_id, name, gender, class_name) for student_id, name, gender, class_name in all_students]
    
    def group_students(self) -> Dict[str, List[Student]]:
        """按照要求进行学生分组：
        - A、B、C、D、E、F共6个组
        - 根据实际学生总数调整各组人数
        - 确保每组男女生比例更均衡
        """
        # 复制学生列表以避免修改原始数据
        students = self.students.copy()
        
        # 按性别分组
        boys = [student for student in students if student.gender == '男']
        girls = [student for student in students if student.gender == '女']
        
        # 计算男女总数
        total_boys = len(boys)
        total_girls = len(girls)
        
        # 随机打乱顺序
        random.shuffle(boys)
        random.shuffle(girls)
        
        # 初始化结果字典
        groups = {group: [] for group in self.group_config}
        
        # 计算每个组理论上应有的男女比例
        boy_ratio = total_boys / self.total_students if self.total_students > 0 else 0
        girl_ratio = total_girls / self.total_students if self.total_students > 0 else 0
        
        # 重置所有学生的分组信息
        for student in students:
            student.group = None
            student.seat_number = None
        
        # 先尝试使用均衡分配算法
        boy_index = 0
        girl_index = 0
        
        # 为每个组分配学生
        for group_name, group_size in self.group_config.items():
            # 计算该组理想的男女生数量
            ideal_boys = max(1, round(group_size * boy_ratio))  # 至少1个男生
            ideal_girls = max(1, round(group_size * girl_ratio))  # 至少1个女生
            
            # 确保理想人数不超过组大小
            if ideal_boys + ideal_girls > group_size:
                if boy_ratio > girl_ratio:
                    ideal_boys = min(ideal_boys, group_size - 1)
                    ideal_girls = group_size - ideal_boys
                else:
                    ideal_girls = min(ideal_girls, group_size - 1)
                    ideal_boys = group_size - ideal_girls
            
            # 分配男生
            assigned_boys = 0
            while assigned_boys < ideal_boys and boy_index < len(boys) and len(groups[group_name]) < group_size:
                student = boys[boy_index]
                student.group = group_name
                groups[group_name].append(student)
                boy_index += 1
                assigned_boys += 1
            
            # 分配女生
            assigned_girls = 0
            while assigned_girls < ideal_girls and girl_index < len(girls) and len(groups[group_name]) < group_size:
                student = girls[girl_index]
                student.group = group_name
                groups[group_name].append(student)
                girl_index += 1
                assigned_girls += 1
            
            # 补充剩余名额
            while len(groups[group_name]) < group_size:
                if boy_index < len(boys):
                    student = boys[boy_index]
                    student.group = group_name
                    groups[group_name].append(student)
                    boy_index += 1
                elif girl_index < len(girls):
                    student = girls[girl_index]
                    student.group = group_name
                    groups[group_name].append(student)
                    girl_index += 1
                else:
                    break
        
        # 检查是否有学生未分配或组人数不符合配置
        total_assigned = sum(len(students) for students in groups.values())
        all_groups_correct = True
        
        for group_name, group_students in groups.items():
            if len(group_students) != self.group_config[group_name]:
                print(f"警告：{group_name}组人数({len(group_students)})不符合配置({self.group_config[group_name]})！")
                all_groups_correct = False
        
        if total_assigned != self.total_students:
            print(f"错误：总分配人数({total_assigned})与总配置人数({self.total_students})不匹配！")
            print("执行强制重新分配...")
        
        # 如果有问题，执行强制重新分配
        if not all_groups_correct or total_assigned != self.total_students:
            # 收集所有学生
            all_students = students.copy()
            
            # 清空所有组
            groups = {group: [] for group in self.group_config}
            
            # 重新随机打乱
            random.shuffle(all_students)
            
            # 按照配置人数重新分配
            current_idx = 0
            for group_name, group_size in self.group_config.items():
                # 确保不会越界
                if current_idx < len(all_students):
                    # 分配学生到组
                    end_idx = min(current_idx + group_size, len(all_students))
                    group_students = all_students[current_idx:end_idx]
                    
                    # 更新学生组信息
                    for student in group_students:
                        student.group = group_name
                        
                    # 添加到组
                    groups[group_name] = group_students
                    current_idx = end_idx
        
        # 为每个组内的学生分配座位号
        for group_name, group_students in groups.items():
            # 随机打乱组内顺序
            random.shuffle(group_students)
            
            # 分配座位号
            for i, student in enumerate(group_students, 1):
                student.seat_number = i
        
        return groups
        
        return groups
    
    def print_group_results(self, groups: Dict[str, List[Student]]):
        """打印分组结果"""
        print("学生分组和座位排号结果：")
        print("=" * 80)
        
        total_boys = 0
        total_girls = 0
        
        for group_name in sorted(groups.keys()):
            students = groups[group_name]
            group_boys = sum(1 for s in students if s.gender == '男')
            group_girls = len(students) - group_boys
            
            total_boys += group_boys
            total_girls += group_girls
            
            print(f"组 {group_name} ({len(students)}人):")
            print(f"  男生: {group_boys}人, 女生: {group_girls}人")
            # 按座位号排序
            students_sorted = sorted(students, key=lambda s: s.seat_number)
            for student in students_sorted:
                print(f"  座位{student.seat_number}: {student.name} ({student.gender}, {student.class_name})")
            print()
        
        print(f"总人数: {self.total_students}人，其中男生{total_boys}人，女生{total_girls}人")
        print("=" * 80)
    
    def export_to_excel(self, groups: Dict[str, List[Student]], filename: str = "学生分组排座结果class1.xlsx"):
        """导出分组结果到Excel文件"""
        # 准备数据
        data = []
        for group, students in groups.items():
            for student in students:
                data.append({
                    "学号": student.student_id,
                    "姓名": student.name,
                    "性别": student.gender,
                    "班级": student.class_name,
                    "组别": student.group,
                    "座位号": student.seat_number
                })
        
        # 创建DataFrame并排序
        df = pd.DataFrame(data)
        df = df.sort_values(by=["组别", "座位号"])
        
        # 导出到Excel
        output_path = f"/Users/xushida/Nutstore Files/我的坚果云/大连交通大学(新)/教学/2025秋-最优化导论/排座程序_python/{filename}"
        df.to_excel(output_path, index=False)
        print(f"分组结果已导出到: {output_path}")

if __name__ == "__main__":
    # 初始化系统，指定Excel文件路径
    excel_file = "/Users/xushida/Nutstore Files/我的坚果云/大连交通大学(新)/教学/2025秋-最优化导论/教学班1(34人)——物流241、242/第一教学班-选课名单.xlsx"
    system = StudentGroupingSystem(excel_file)
    
    # 进行分组
    groups = system.group_students()
    
    # 打印结果
    system.print_group_results(groups)
    
    # 导出到Excel
    system.export_to_excel(groups)