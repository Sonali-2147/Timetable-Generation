import pandas as pd
import random

# Define subjects and their corresponding teachers
subjects_and_teachers = {
    "SAP/Career Counselling": ["Vanita J"],
    "Life Skills": ["Vanita J"],
    "Art": ["HR", "Pankaj", "Shridevi"],
    "Clay Modelling": ["Shridevi"],
    "Music": ["Sanjukta", "Shlok"],
    "Dance": ["Pooja B"],
    "Library": ["Pooja S"],
    "DEED": ["Isha"],
    "Assembly": ["Neesha.R", "Shilpa.K", "New", "Shubhi", "Priya"],
    "Wonder Time": ["Neesha.R", "Shilpa.K", "New", "Shubhi", "Priya"],
    "SUPW": ["New Sports Tr"],
    "Research and Referral": ["New Sports Tr"],
    "Yoga": ["New Sports Tr"],
    "Karate": ["RAJU"]
}

# Define grades, divisions, days, and timings
grades = ["Jr.Kg", "Sr.Kg", "1st"]
divisions = ["A", "B", "C", "D"]
days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
timings = [
    "08:00 - 08:15 (Home Room)",
    "08:15 - 08:50 (1st)",
    "08:50 - 09:25 (2nd)",
    "09:25 - 09:45 (Break)",
    "09:45 - 10:15 (3rd)",
    "10:15 - 10:45 (4th)",
    "10:45 - 11:15 (5th)",
    "11:15 - 11:45 (6th)",
    "11:45 - 12:15 (7th)",
    "12:15 - 12:45 (8th)",
    "12:45 - 01:15 (Break)",
    "01:15 - 01:45 (9th)",
    "01:45 - 02:15 (10th)",
    "02:15 - 02:20 (Dispersal)"
]

# Weekly subject-hour requirements
subject_hours = {
    "Jr.Kg": {
        "English": 5,
        "Hindi": 4,
        "Mathematics": 4,
        "PE/CCA Sports": 2,
        "EVS": 2,
        "Art": 2,
        "Clay Modelling": 1,
        "Music": 1,
        "Dance": 1,
        "Yoga": 0.5,
        "Library": 1,
        "Wonder Time": 1,
        "Circle Time": 0.5,
        "Life Skill": 0.5,
        "DEED": 1,
        "Environmental Education (EE)": 2,
        "Play Pen": 0.5
    },
    "Sr.Kg": {
        "English": 5,
        "Hindi": 4,
        "Mathematics": 4.5,
        "PE/CCA Sports": 2,
        "EVS": 3,
        "Art": 2,
        "Clay Modelling": 1,
        "Music": 1,
        "Dance": 1,
        "Yoga": 0.5,
        "Library": 1,
        "Wonder Time": 1,
        "Circle Time": 0.5,
        "Life Skill": 0.5,
        "DEED": 0.5,
        "Environmental Education (EE)": 1
    },
    "1st": {
        "English": 8,
        "Hindi": 6,
        "Mathematics": 7,
        "PE/CCA Sports": 2,
        "EVS": 5,
        "Marathi": 2,
        "Computer Studies": 2,
        "Art": 2,
        "Clay Modelling": 1,
        "Music": 1,
        "Dance": 1,
        "Yoga": 1,
        "Library": 1,
        "Wonder Time": 1,
        "Circle Time": 1,
        "Life Skill": 1,
        "DEED": 1,
        "Assembly": 1,
        "Environmental Education (EE)": 1,
        "Karate": 1,
        "Class Tests": 1
    }
}

# Assign teachers to subjects
teachers = {subject: ', '.join(teacher_list) for subject, teacher_list in subjects_and_teachers.items()}

home_teachers = {
    "Jr.Kg": {"A": "Kim, Rukhsaar", "B": "Maria, Shital", "C": "Shreen, Jijikar", "D": "Priyanka, Darshana"},
    "Sr.Kg": {"A": "Pooja, Nupur", "B": "Lakshmipriya, Jyoti", "C": "Shubha, Priyanka", "D": "Ankita, Bhakti"},
    "1st": {"A": "Remya, Saraswati", "B": "Kanak, Matilda", "C": "Rishita, Shashwati", "D": "Charu, Maariyah"}
}

# Generate a timetable for a grade and division
def generate_timetable(grade, division):
    weekly_subject_hours = subject_hours[grade].copy()
    timetable = {day: [None for _ in timings] for day in days}

    # Generate the timetable
    for day in days:
        day_subjects = [subject for subject, hours in weekly_subject_hours.items() if hours > 0 and subject in teachers]
        random.shuffle(day_subjects)

        subject_index = 0
        for i, timing in enumerate(timings):
            if "Break" in timing:
                timetable[day][i] = "Break"
            elif "Dispersal" in timing:
                timetable[day][i] = "Dispersal"
            else:
                if subject_index < len(day_subjects):
                    subject = day_subjects[subject_index]
                    timetable[day][i] = f"{subject} ({teachers[subject]})"
                    weekly_subject_hours[subject] -= 1

                    if weekly_subject_hours[subject] == 0:
                        subject_index += 1
                

    for day in days:
        timetable[day][0] = f"Home Room ({home_teachers[grade][division]})"

    return timetable

# Generate and save all timetables
all_timetables = {}
for grade in grades:
    all_timetables[grade] = {}
    for division in divisions:
        all_timetables[grade][division] = generate_timetable(grade, division)

with pd.ExcelWriter("Timetables.xlsx", engine="openpyxl") as writer:
    for grade, divisions in all_timetables.items():
        for division, timetable in divisions.items():
            df = pd.DataFrame.from_dict(timetable, orient="index", columns=timings)
            sheet_name = f"{grade}_{division}"
            df.to_excel(writer, sheet_name=sheet_name)

    # Home room data
    data = {
        "Class": [],
        "Home Room Teachers": [],
        "Subjects": [],
        "Class Allotted": []
    }
    for grade, div_teachers in home_teachers.items():
        for division, teacher in div_teachers.items():
            data["Class"].append(f"{grade} {division}")
            data["Home Room Teachers"].append(teacher)
            data["Subjects"].append("All")
            data["Class Allotted"].append("Home Room")

    home_room_df = pd.DataFrame(data)
    home_room_df.to_excel(writer, sheet_name="Home Room", index=False)

print("Timetables successfully saved to 'Timetables.xlsx'.")
