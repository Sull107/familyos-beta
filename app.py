import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import date, datetime

# ==================================================
# LOGIN SETTINGS
# ==================================================

USERS_FILE = Path(".") / "users.xlsx"

default_users = pd.DataFrame({
    "Username": ["nick", "matthew"],
    "Password": ["familyos", "familyos"],
    "Role": ["Admin", "Child"],
    "Display Name": ["Nick", "Matthew"]
})

if USERS_FILE.exists():
    users = pd.read_excel(USERS_FILE)
else:
    users = default_users.copy()
    users.to_excel(USERS_FILE, index=False)

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if "username" not in st.session_state:
    st.session_state.username = ""

if "role" not in st.session_state:
    st.session_state.role = ""

if "display_name" not in st.session_state:
    st.session_state.display_name = ""

# ==================================================
# 1. PAGE SETUP
# ==================================================

st.set_page_config(
    page_title="FamilyOS",
    page_icon="🏠",
    layout="wide"
)

# ==================================================
# LOGIN SCREEN
# ==================================================

if not st.session_state.logged_in:

    st.title("🏠 FamilyOS")

    st.subheader("FamilyOS Beta Login")

    username = st.text_input(
        "Username"
    )

    password = st.text_input(
        "Password",
        type="password"
    )

    if st.button("Login"):

        matched_user = users[
            (users["Username"].astype(str).str.lower() == username.lower()) &
            (users["Password"].astype(str) == password)
        ]

        if not matched_user.empty:

            user_row = matched_user.iloc[0]

            st.session_state.logged_in = True
            st.session_state.username = user_row["Username"]
            st.session_state.role = user_row["Role"]
            st.session_state.display_name = user_row["Display Name"]

            st.rerun()

        else:
            st.error("Incorrect username or password.")

    st.stop()

# ==================================================
# 2. FILE PATHS
# ==================================================

DATA_DIR = Path(".")
EVENTS_FILE = DATA_DIR / "events.xlsx"
TASKS_FILE = DATA_DIR / "tasks.xlsx"
GOALS_FILE = DATA_DIR / "goals.xlsx"
CHAT_FILE = DATA_DIR / "chat.xlsx"
FAMILY_FILE = DATA_DIR / "family_members.xlsx"
SHOPPING_FILE = DATA_DIR / "shopping.xlsx"
MEALS_FILE = DATA_DIR / "meals.xlsx"
BUDGET_FILE = DATA_DIR / "budget.xlsx"
PETS_FILE = DATA_DIR / "pets.xlsx"
HOUSEHOLD_FILE = DATA_DIR / "household.xlsx"

# ==================================================
# 3. DEFAULT DATA
# ==================================================

default_events = pd.DataFrame({
    "Date": [str(date.today()), str(date.today())],
    "Time": ["18:00", "20:00"],
    "Event": ["Matthew football training", "Bins out"],
    "Category": ["Sports", "Household"],
    "Owner": ["Matthew", "Dad"]
})

default_tasks = pd.DataFrame({
    "Task": ["Feed Bailey", "Homework", "Empty bins"],
    "Assigned To": ["Matthew", "Matthew", "Dad"],
    "Due Date": [str(date.today()), str(date.today()), str(date.today())],
    "Points": [10, 20, 15],
    "Status": ["Open", "Open", "Open"]
})

default_goals = pd.DataFrame({
    "Goal": ["Florida Holiday", "Matthew Gaming Fund"],
    "Target (£)": [5000.00, 600.00],
    "Saved (£)": [1200.00, 150.00],
    "Owner": ["Family", "Matthew"]
})

default_chat = pd.DataFrame({
    "Time": [datetime.now().strftime("%Y-%m-%d %H:%M")],
    "Person": ["Betty"],
    "Message": ["Welcome to FamilyOS. I’ll help keep the family organised."]
})

default_family = pd.DataFrame({
    "Name": ["Nick", "Matthew", "Bailey"],
    "Relationship": ["Dad", "Son", "Dog"],
    "Birthday": ["1985-01-01", "2014-01-01", "2020-01-01"]
})

default_shopping = pd.DataFrame({
    "Item": ["Milk", "Bread", "Dog Food"],
    "Category": ["Groceries", "Groceries", "Pets"],
    "Priority": ["High", "Medium", "High"],
    "Added By": ["Nick", "Matthew", "Betty"],
    "Status": ["Open", "Open", "Open"]
})

default_meals = pd.DataFrame({
    "Meal": ["Chicken Fajitas", "Spaghetti Bolognese", "Homemade Pizza"],
    "Meal Type": ["Dinner", "Dinner", "Dinner"],
    "Ingredients": [
        "Chicken, wraps, peppers, cheese, salsa",
        "Mince, pasta, tomato sauce, onions, garlic",
        "Pizza bases, cheese, pepperoni, tomato sauce"
    ],
    "Suggested By": ["Matthew", "Nick", "Family"],
    "Difficulty": ["Easy", "Medium", "Easy"]
})

default_budget = pd.DataFrame({
    "Date": [str(date.today()), str(date.today()), str(date.today())],
    "Category": ["Food", "Fuel", "Pets"],
    "Description": ["Tesco shop", "Petrol", "Dog food"],
    "Amount (£)": [75.00, 50.00, 35.00]
})

default_pets = pd.DataFrame({
    "Pet Name": ["Bailey"],
    "Animal": ["Dog"],
    "Breed": ["Unknown"],
    "Birthday": ["2020-01-01"],
    "Vet": ["Local Vet"],
    "Insurance Renewal": ["2026-12-31"],
    "Next Vaccination": [str(date.today())],
    "Medication": ["None"],
    "Weight (kg)": [0.0],
    "Notes": ["Good boy."]
})

default_household = pd.DataFrame({
    "Item": [
        "Boiler Service",
        "Home Insurance",
        "Energy Tariff Review",
        "Car MOT",
        "Car Insurance"
    ],
    "Category": [
        "Home",
        "Home",
        "Home",
        "Vehicle",
        "Vehicle"
    ],
    "Due Date": [
        "2026-12-01",
        "2026-11-01",
        "2026-09-01",
        "2026-08-01",
        "2026-10-01"
    ],
    "Owner": [
        "Nick",
        "Nick",
        "Nick",
        "Nick",
        "Nick"
    ],
    "Status": [
        "Open",
        "Open",
        "Open",
        "Open",
        "Open"
    ],
    "Notes": [
        "Annual boiler service",
        "Review renewal price",
        "Check fixed tariff options",
        "Book MOT before expiry",
        "Compare insurance quotes"
    ]
})

budget_categories = pd.DataFrame({
    "Category": ["Food", "Fuel", "Entertainment", "Holidays", "Pets", "Household", "Other"],
    "Monthly Budget (£)": [400.00, 200.00, 150.00, 300.00, 100.00, 200.00, 100.00]
})

# ==================================================
# 4. LOAD / SAVE FUNCTIONS
# ==================================================

def load_data(file_path, default_df):
    if file_path.exists():
        try:
            return pd.read_excel(file_path)
        except Exception:
            return default_df.copy()
    default_df.to_excel(file_path, index=False)
    return default_df.copy()


def save_data(df, file_path):
    df.to_excel(file_path, index=False)


events = load_data(EVENTS_FILE, default_events) 
tasks = load_data(TASKS_FILE, default_tasks) 
goals = load_data(GOALS_FILE, default_goals) 
chat = load_data(CHAT_FILE, default_chat) 
family = load_data(FAMILY_FILE, default_family) 
shopping = load_data(SHOPPING_FILE, default_shopping) 
meals = load_data(MEALS_FILE, default_meals)
budget = load_data(BUDGET_FILE, default_budget)
pets = load_data(PETS_FILE, default_pets)
household = load_data(HOUSEHOLD_FILE, default_household)

# ==================================================
# 5. CALCULATIONS
# ==================================================

today = str(date.today())
today_date = date.today()

today_events = events[events["Date"].astype(str) == today] if not events.empty else pd.DataFrame() 
open_tasks = tasks[tasks["Status"] == "Open"] if not tasks.empty else pd.DataFrame() 
completed_tasks = tasks[tasks["Status"] == "Done"] if not tasks.empty else pd.DataFrame()

open_shopping = shopping[shopping["Status"] == "Open"] if not shopping.empty else pd.DataFrame() 
high_priority_shopping = open_shopping[open_shopping["Priority"] == "High"] if not open_shopping.empty else pd.DataFrame()

total_tasks = len(tasks)
done_tasks = len(completed_tasks)
task_score = round((done_tasks / total_tasks) * 100, 0) if total_tasks > 0 else 100

goal_progress_avg = 0
if not goals.empty:
    goals["Progress %"] = (goals["Saved (£)"] / goals["Target (£)"] * 100).clip(upper=100)
    goal_progress_avg = goals["Progress %"].mean()

birthday_reminders = pd.DataFrame()

if not family.empty:
    family_calc = family.copy()
    family_calc["Birthday"] = pd.to_datetime(family_calc["Birthday"], errors="coerce")

    ages = []
    days_until = []
    next_ages = []

    for _, row in family_calc.iterrows():
        birthday = row["Birthday"]

        if pd.isna(birthday):
            ages.append("Unknown")
            days_until.append("Unknown")
            next_ages.append("Unknown")
        else:
            birthday_date = birthday.date()
            age = today_date.year - birthday_date.year

            if (today_date.month, today_date.day) < (birthday_date.month, birthday_date.day):
                age -= 1

            next_birthday = date(today_date.year, birthday_date.month, birthday_date.day)

            if next_birthday < today_date:
                next_birthday = date(today_date.year + 1, birthday_date.month, birthday_date.day)

            days = (next_birthday - today_date).days

            ages.append(age)
            days_until.append(days)
            next_ages.append(age + 1)

    family_calc["Age"] = ages
    family_calc["Days Until Birthday"] = days_until
    family_calc["Turning"] = next_ages

    birthday_reminders = family_calc[
        family_calc["Days Until Birthday"].apply(lambda x: isinstance(x, int) and x <= 30)
    ]

monthly_budget_total = budget_categories["Monthly Budget (£)"].sum() 
monthly_spend_total = budget["Amount (£)"].sum() if not budget.empty else 0 
remaining_budget = monthly_budget_total - monthly_spend_total 
budget_used_pct = (monthly_spend_total / monthly_budget_total * 100) if monthly_budget_total > 0 else 0

# Household calculations

household_reminders = pd.DataFrame()

if not household.empty:

    household_calc = household.copy()
    household_calc["Due Date"] = pd.to_datetime(
        household_calc["Due Date"],
        errors="coerce"
    )

    days_until_due = []

    for _, row in household_calc.iterrows():

        due_date = row["Due Date"]

        if pd.isna(due_date):
            days_until_due.append("Unknown")
        else:
            days_until_due.append(
                (due_date.date() - today_date).days
            )

    household_calc["Days Until Due"] = days_until_due

    household_reminders = household_calc[
        household_calc["Days Until Due"].apply(
            lambda x: isinstance(x, int) and x <= 30
        )
    ]

calendar_readiness = 100 if len(today_events) <= 2 else 75 
task_readiness = 100 - min(len(open_tasks) * 15, 70) 
goal_readiness = min(goal_progress_avg, 100) 
birthday_readiness = 75 if not birthday_reminders.empty else 100 
shopping_readiness = 100 - min(len(open_shopping) * 10, 60) 
budget_readiness = max(0, 100 - max(0, budget_used_pct - 80))
household_readiness = 75 if not household_reminders.empty else 100

readiness_score = round(
    (
        (calendar_readiness * 0.15) +
        (task_readiness * 0.22) +
        (goal_readiness * 0.15) +
        (birthday_readiness * 0.10) +
        (shopping_readiness * 0.10) +
        (budget_readiness * 0.13) +
        (household_readiness * 0.15)
    ),
    0
)


if readiness_score >= 80:
    readiness_status = "🟢 Ready"
elif readiness_score >= 60:
    readiness_status = "🟠 Needs Focus"
else:
    readiness_status = "🔴 Busy / At Risk"

# ==================================================
# 6. STYLE
# ==================================================

st.markdown(
    """
    <style>
    .kpi-card {
        background-color: white;
        padding: 20px;
        border-radius: 14px;
        box-shadow: 0 1px 8px rgba(0,0,0,0.08);
        border: 1px solid #e6e6e6;
    }
    .kpi-title {
        font-size: 14px;
        color: #6b7280;
    }
    .kpi-value {
        font-size: 30px;
        font-weight: 700;
        color: #111827;
    }
    .kpi-caption {
        font-size: 13px;
        color: #6b7280;
    }
    </style>
    """,
    unsafe_allow_html=True
)


def kpi_card(title, value, caption):
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-title">{title}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-caption">{caption}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

# ==================================================
# LOGOUT BAR
# ==================================================

logout_col1, logout_col2 = st.columns([4, 1])

with logout_col1:
    st.caption(
        f"Logged in as: {st.session_state.display_name} "
        f"({st.session_state.role})"
    )

with logout_col2:
    logout = st.button(
        "🚪 Logout",
        use_container_width=True
    )

    if logout:
        st.session_state.logged_in = False
        st.session_state.username = ""
        st.session_state.role = ""
        st.session_state.display_name = ""
        st.rerun()

# ==================================================
# 7. HEADER / NAVIGATION
# ==================================================

st.title("🏠 FamilyOS")
st.caption("Organise. Connect. Simplify.")

current_role = st.session_state.role
is_admin = current_role == "Admin"
is_parent = current_role in ["Admin", "Parent"] 
is_child = current_role == "Child"

tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9, tab10, tab11, tab12, tab13 = st.tabs([
      "🏠 Home",
      "📅 Calendar",
      "✅ Tasks",
      "🎯 Goals",
      "💬 Family Chat",
      "🤖 BettyAI",
      "👨‍👩‍👧 Family Members",
      "🛒 Shopping",
      "💰 Budget",
      "🐶 Pets",
      "🏡 Household",
      "🔔 Notifications",
      "💾 Backup"
 ])


# ==================================================
# 8. HOME DASHBOARD
# ==================================================

with tab1:

    st.header("🏠 Family Dashboard")

    current_hour = datetime.now().hour

    if current_hour < 12:
        greeting = "Good Morning"
    elif current_hour < 18:
        greeting = "Good Afternoon"
    else:
        greeting = "Good Evening"

    st.subheader(f"{greeting}, {st.session_state.display_name} 👋")

    # ==================================================
    # BETTY DAILY BRIEFING
    # ==================================================

    betty_summary = []

    if len(today_events) > 0:
        betty_summary.append(f"📅 {len(today_events)} event(s) today")

    if len(open_tasks) > 0:
        betty_summary.append(f"✅ {len(open_tasks)} open task(s)")

    if len(open_shopping) > 0:
        betty_summary.append(f"🛒 {len(open_shopping)} shopping item(s)")

    if not birthday_reminders.empty:
        betty_summary.append(f"🎂 {len(birthday_reminders)} birthday reminder(s)")

    if "household_reminders" in globals() and not household_reminders.empty:
        betty_summary.append(f"🏡 {len(household_reminders)} household reminder(s)")

    if len(betty_summary) > 0:
        briefing_text = " | ".join(betty_summary)
    else:
        briefing_text = "✅ Nothing urgent showing today."

    st.info(
        f"""
🤖 **Betty Daily Briefing**

Family Readiness: **{readiness_score:.0f}/100** {readiness_status}

{briefing_text}
"""
    )

    # ==================================================
    # BETTY RECOMMENDATION OF THE DAY
    # ==================================================

    if len(open_tasks) >= 3:
        recommendation = "Complete outstanding tasks"
        reason = f"There are currently {len(open_tasks)} open tasks reducing family readiness."

    elif len(open_shopping) >= 3:
        recommendation = "Complete the shopping list"
        reason = f"There are currently {len(open_shopping)} shopping items waiting."

    elif "household_reminders" in globals() and not household_reminders.empty:
        recommendation = "Review household maintenance"
        reason = f"There are {len(household_reminders)} household items due soon."

    elif not birthday_reminders.empty:
        next_birthday = birthday_reminders.sort_values("Days Until Birthday").iloc[0]
        recommendation = f"Plan for {next_birthday['Name']}'s birthday"
        reason = f"Birthday is in {next_birthday['Days Until Birthday']} days."

    else:
        recommendation = "Enjoy your day"
        reason = "Family life is currently under control."

    st.success(
        f"""
🎯 **Betty Recommendation of the Day**

**{recommendation}**

{reason}
"""
    )

    # ==================================================
    # TODAY'S FOCUS
    # ==================================================

    focus_items = []

    if len(today_events) > 0:
        next_event = today_events.iloc[0]["Event"]
        focus_items.append(f"Prepare for **{next_event}**.")

    if len(open_shopping) > 0:
        focus_items.append(f"Complete **{len(open_shopping)} shopping item(s)**.")

    if len(open_tasks) > 0:
        focus_items.append(f"Complete at least **1 of {len(open_tasks)} open task(s)**.")

    if not birthday_reminders.empty:
        next_birthday = birthday_reminders.sort_values("Days Until Birthday").iloc[0]
        focus_items.append(f"Plan for **{next_birthday['Name']}**'s birthday.")

    if "household_reminders" in globals() and not household_reminders.empty:
        focus_items.append(f"Review **{len(household_reminders)} household reminder(s)**.")

    if len(focus_items) > 0:
        focus_text = "\n\n".join([f"• {item}" for item in focus_items])
    else:
        focus_text = "Family life is under control today. Keep FamilyOS updated."

    st.warning(
        f"""
🎯 **Today's Focus**

{focus_text}
"""
    )

    # ==================================================
    # KPI CARDS
    # ==================================================

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        kpi_card("Family Readiness", f"{readiness_score:.0f}/100", readiness_status)

    with col2:
        kpi_card("Today’s Events", len(today_events), "Events happening today")

    with col3:
        kpi_card("Open Tasks", len(open_tasks), "Tasks needing completed")

    with col4:
        kpi_card("Budget Remaining", f"£{remaining_budget:,.0f}", "This month")

    st.divider()

    # ==================================================
    # FAMILY READINESS BREAKDOWN
    # ==================================================

    st.subheader("📊 Family Readiness Breakdown")

    bcol1, bcol2, bcol3, bcol4, bcol5, bcol6, bcol7 = st.columns(7)

    with bcol1:
        st.write("Calendar")
        st.progress(calendar_readiness / 100)
        st.caption(f"{calendar_readiness:.0f}%")

    with bcol2:
        st.write("Tasks")
        st.progress(task_readiness / 100)
        st.caption(f"{task_readiness:.0f}%")

    with bcol3:
        st.write("Goals")
        st.progress(goal_readiness / 100)
        st.caption(f"{goal_readiness:.0f}%")

    with bcol4:
        st.write("Birthdays")
        st.progress(birthday_readiness / 100)
        st.caption(f"{birthday_readiness:.0f}%")

    with bcol5:
        st.write("Shopping")
        st.progress(shopping_readiness / 100)
        st.caption(f"{shopping_readiness:.0f}%")

    with bcol6:
        st.write("Budget")
        st.progress(budget_readiness / 100)
        st.caption(f"{budget_readiness:.0f}%")

    with bcol7:
        st.write("Household")
        st.progress(household_readiness / 100)
        st.caption(f"{household_readiness:.0f}%")

# ==================================================
# 9. FAMILY CALENDAR
# ==================================================

with tab2:

    st.header("📅 Family Calendar")

    st.dataframe(events, use_container_width=True)

    st.subheader("➕ Add Family Event")

    with st.form("add_event"):
        event_date = st.date_input("Date", value=date.today())
        event_time = st.text_input("Time", value="18:00")
        event_name = st.text_input("Event")
        event_category = st.selectbox(
            "Category",
            ["School", "Work", "Sports", "Family", "Holiday", "Pet", "Household", "Other"]
        )
        event_owner = st.text_input("Owner", value="Family")

        submitted = st.form_submit_button("Add Event")

        if submitted:
            new_event = pd.DataFrame([{
                "Date": str(event_date),
                "Time": event_time,
                "Event": event_name,
                "Category": event_category,
                "Owner": event_owner
            }])

            events = pd.concat([events, new_event], ignore_index=True)
            save_data(events, EVENTS_FILE)
            st.success("Event added.")
            st.rerun()

# ==================================================
# 10. TASKS & CHORES
# ==================================================

with tab3:

    st.header("✅ Tasks & Chores")

    st.dataframe(tasks, use_container_width=True)

    st.subheader("➕ Add Task / Chore")

    with st.form("add_task"):
        task_name = st.text_input("Task")
        assigned_to = st.text_input("Assigned To")
        due_date = st.date_input("Due Date", value=date.today())
        points = st.number_input("Points", min_value=0, value=10, step=5)
        status = st.selectbox("Status", ["Open", "Done"])

        submitted = st.form_submit_button("Add Task")

        if submitted:
            new_task = pd.DataFrame([{
                "Task": task_name,
                "Assigned To": assigned_to,
                "Due Date": str(due_date),
                "Points": points,
                "Status": status
            }])

            tasks = pd.concat([tasks, new_task], ignore_index=True)
            save_data(tasks, TASKS_FILE)
            st.success("Task added.")
            st.rerun()

    st.divider()

    st.subheader("🏆 Points Leaderboard")

    if not tasks.empty:
        points_df = tasks[tasks["Status"] == "Done"].groupby("Assigned To")["Points"].sum().reset_index()
        points_df = points_df.sort_values("Points", ascending=False)
        st.dataframe(points_df, use_container_width=True)

# ==================================================
# 11. FAMILY GOALS
# ==================================================

with tab4:

    st.header("🎯 Family Goals")

    if not goals.empty:
        goals_display = goals.copy()
        goals_display["Progress %"] = (goals_display["Saved (£)"] / goals_display["Target (£)"] * 100).round(1)
        st.dataframe(goals_display, use_container_width=True)

        for _, row in goals_display.iterrows():
            st.write(f"**{row['Goal']}**")
            st.progress(min(row["Progress %"] / 100, 1.0))
            st.caption(f"£{row['Saved (£)']:,.2f} saved of £{row['Target (£)']:,.2f} target")

    st.subheader("➕ Add Family Goal")

    with st.form("add_goal"):
        goal_name = st.text_input("Goal")
        target_amount = st.number_input("Target (£)", min_value=0.0, value=1000.0, step=100.0)
        saved_amount = st.number_input("Saved (£)", min_value=0.0, value=0.0, step=50.0)
        owner = st.text_input("Owner", value="Family")

        submitted = st.form_submit_button("Add Goal")

        if submitted:
            new_goal = pd.DataFrame([{
                "Goal": goal_name,
                "Target (£)": target_amount,
                "Saved (£)": saved_amount,
                "Owner": owner
            }])

            goals = pd.concat([goals, new_goal], ignore_index=True)
            save_data(goals, GOALS_FILE)
            st.success("Goal added.")
            st.rerun()

# ==================================================
# 12. FAMILY CHAT
# ==================================================

with tab5:

    st.header("💬 Family Chat")

    st.dataframe(chat, use_container_width=True)

    st.subheader("➕ Add Message")

    with st.form("add_chat"):
        person = st.text_input("Person", value="Nick")
        message = st.text_area("Message")

        submitted = st.form_submit_button("Send Message")

        if submitted:
            new_message = pd.DataFrame([{
                "Time": datetime.now().strftime("%Y-%m-%d %H:%M"),
                "Person": person,
                "Message": message
            }])

            chat = pd.concat([chat, new_message], ignore_index=True)
            save_data(chat, CHAT_FILE)
            st.success("Message added.")
            st.rerun()

# ==================================================
# 13. BETTYAI COMMAND CENTRE 3.0
# ==================================================

with tab6:

    st.header("🤖 BettyAI Command Centre 3.0")

    st.write("Betty reviews your family data and highlights what needs attention, why it matters, and what to do next.")

    st.divider()

    # ==================================================
    # BETTY OVERVIEW
    # ==================================================

    st.subheader("🏠 Family Readiness Review")

    rcol1, rcol2, rcol3, rcol4 = st.columns(4)

    with rcol1:
        st.metric("Readiness Score", f"{readiness_score:.0f}/100", readiness_status)

    with rcol2:
        st.metric("Open Tasks", len(open_tasks))

    with rcol3:
        st.metric("Shopping Items", len(open_shopping))

    with rcol4:
        st.metric("Budget Used", f"{budget_used_pct:.1f}%")

    st.divider()

    # ==================================================
    # BETTY INSIGHTS ENGINE
    # ==================================================

    st.subheader("🧠 Betty Insights Engine")

    insight_rows = []

    readiness_map = {
        "Calendar": calendar_readiness,
        "Tasks": task_readiness,
        "Goals": goal_readiness,
        "Birthdays": birthday_readiness,
        "Shopping": shopping_readiness,
        "Budget": budget_readiness,
        "Household": household_readiness
    }

    lowest_area = min(readiness_map, key=readiness_map.get)
    lowest_score = readiness_map[lowest_area]

    if lowest_score < 80:
        insight_rows.append({
            "Area": lowest_area,
            "Insight": f"{lowest_area} is currently the lowest readiness area.",
            "Impact": f"Score: {lowest_score:.0f}/100",
            "Betty Action": f"Focus on improving {lowest_area.lower()} today."
        })

    if len(open_tasks) > 0:
        insight_rows.append({
            "Area": "Tasks",
            "Insight": f"There are {len(open_tasks)} open task(s).",
            "Impact": "Open tasks reduce family readiness.",
            "Betty Action": f"Start with: {open_tasks.iloc[0]['Task']}"
        })

    if len(open_shopping) > 0:
        insight_rows.append({
            "Area": "Shopping",
            "Insight": f"There are {len(open_shopping)} shopping item(s) still open.",
            "Impact": "Open shopping items create household friction.",
            "Betty Action": "Clear high priority shopping items first."
        })

    if budget_used_pct >= 80:
        insight_rows.append({
            "Area": "Budget",
            "Insight": f"Budget is {budget_used_pct:.1f}% used.",
            "Impact": "Spending is getting close to the monthly limit.",
            "Betty Action": "Review non-essential spending."
        })

    if "household_reminders" in globals() and not household_reminders.empty:
        insight_rows.append({
            "Area": "Household",
            "Insight": f"There are {len(household_reminders)} household reminder(s).",
            "Impact": "Upcoming renewals or maintenance need attention.",
            "Betty Action": "Review the Household tab."
        })

    if not birthday_reminders.empty:
        next_birthday = birthday_reminders.sort_values("Days Until Birthday").iloc[0]

        insight_rows.append({
            "Area": "Birthdays",
            "Insight": f"{next_birthday['Name']}'s birthday is coming up.",
            "Impact": f"Due in {next_birthday['Days Until Birthday']} day(s).",
            "Betty Action": "Plan card, gift or family activity."
        })

    if "pets" in globals() and not pets.empty:
        insight_rows.append({
            "Area": "Pets",
            "Insight": "Pet records are active in FamilyOS.",
            "Impact": "Betty can track vaccinations, insurance and care.",
            "Betty Action": "Review Pet Centre regularly."
        })

    if insight_rows:
        insights_df = pd.DataFrame(insight_rows)
        st.dataframe(insights_df, use_container_width=True)
    else:
        st.success("✅ Betty has not found any major issues. Family life looks under control.")

    st.divider()

    # ==================================================
    # BETTY DAILY PRIORITIES
    # ==================================================

    st.subheader("📌 Betty Daily Priorities")

    priorities = []

    if len(open_tasks) > 0:
        priorities.append({
            "Priority": "Task",
            "Action": f"Complete or review: {open_tasks.iloc[0]['Task']}",
            "Reason": "This will improve task readiness."
        })

    if len(today_events) > 0:
        priorities.append({
            "Priority": "Calendar",
            "Action": f"Prepare for: {today_events.iloc[0]['Event']} at {today_events.iloc[0]['Time']}",
            "Reason": "This is the next visible event today."
        })

    if len(high_priority_shopping) > 0:
        priorities.append({
            "Priority": "Shopping",
            "Action": f"Buy high priority item: {high_priority_shopping.iloc[0]['Item']}",
            "Reason": "High priority shopping items need attention."
        })

    if budget_used_pct > 80:
        priorities.append({
            "Priority": "Budget",
            "Action": "Review monthly spending",
            "Reason": f"Budget is already {budget_used_pct:.1f}% used."
        })

    if "household_reminders" in globals() and not household_reminders.empty:
        priorities.append({
            "Priority": "Household",
            "Action": "Review household maintenance items",
            "Reason": "Some renewals or services are due soon."
        })

    if priorities:
        priorities_df = pd.DataFrame(priorities).head(5)
        st.dataframe(priorities_df, use_container_width=True)

        st.warning(
            f"🎯 Betty recommends starting with: **{priorities_df.iloc[0]['Action']}**"
        )
    else:
        st.success("✅ No urgent daily priorities detected.")

    st.divider()

    # ==================================================
    # BETTY ACTION PLAN
    # ==================================================

    st.subheader("🎯 Betty Action Plan")

    action_plan = []

    if len(open_tasks) > 0:
        action_plan.append(f"Complete **{open_tasks.iloc[0]['Task']}** or reassign it.")

    if len(open_shopping) > 0:
        action_plan.append("Clear the shopping list, starting with high priority items.")

    if budget_used_pct > 80:
        action_plan.append("Review budget spend before adding more expenses.")

    if "household_reminders" in globals() and not household_reminders.empty:
        action_plan.append("Review household renewals or maintenance due soon.")

    if not birthday_reminders.empty:
        next_birthday = birthday_reminders.sort_values("Days Until Birthday").iloc[0]
        action_plan.append(f"Start planning for **{next_birthday['Name']}**'s birthday.")

    if not action_plan:
        action_plan.append("No urgent action required. Keep FamilyOS updated.")

    for item in action_plan:
        st.write(f"• {item}")

    st.divider()

    # ==================================================
    # ASK BETTY
    # ==================================================

    st.subheader("💬 Ask Betty")

    st.write("Try asking about today, focus, tasks, budget, shopping, birthdays, household, Bailey or readiness.")

    user_question = st.text_input("Ask Betty something about the family")

    if user_question:

        question = user_question.lower()

        if "focus" in question or "priority" in question:
            if priorities:
                st.info(f"Betty recommends focusing on: **{priorities[0]['Action']}**")
            else:
                st.info("Betty does not see any urgent priorities right now.")

        elif "insight" in question or "why" in question:
            if insight_rows:
                st.info(
                    f"The biggest readiness issue is currently **{lowest_area}** "
                    f"with a score of **{lowest_score:.0f}/100**."
                )
            else:
                st.info("Betty has not found any major readiness issues.")

        elif "budget" in question or "money" in question:
            st.info(
                f"Monthly budget is £{monthly_budget_total:,.2f}. "
                f"Spend is £{monthly_spend_total:,.2f}. "
                f"Remaining budget is £{remaining_budget:,.2f}. "
                f"Budget used is {budget_used_pct:.1f}%."
            )

        elif "shopping" in question:
            st.info(
                f"There are {len(open_shopping)} open shopping item(s), "
                f"including {len(high_priority_shopping)} high priority item(s)."
            )

        elif "birthday" in question:
            if not birthday_reminders.empty:
                next_birthday = birthday_reminders.sort_values("Days Until Birthday").iloc[0]
                st.info(
                    f"The next birthday is **{next_birthday['Name']}** in "
                    f"{next_birthday['Days Until Birthday']} day(s)."
                )
            else:
                st.info("There are no birthdays due in the next 30 days.")

        elif "task" in question or "chore" in question:
            if len(open_tasks) > 0:
                st.info(
                    f"There are currently {len(open_tasks)} open task(s). "
                    f"First one to review: **{open_tasks.iloc[0]['Task']}**."
                )
            else:
                st.info("There are no open tasks right now.")

        elif "today" in question:
            st.info(
                f"Today you have {len(today_events)} event(s), {len(open_tasks)} open task(s), "
                f"{len(open_shopping)} shopping item(s), and readiness is {readiness_score:.0f}/100."
            )

        elif "household" in question:
            if "household_reminders" in globals() and not household_reminders.empty:
                st.info(f"There are {len(household_reminders)} household reminder(s) due soon.")
            else:
                st.info("There are no urgent household reminders.")

        elif "bailey" in question or "pet" in question:
            st.info("Check the Pet Centre for vet, vaccination, insurance and medication details.")

        elif "health" in question or "readiness" in question:
            st.info(
                f"Family Readiness is **{readiness_score:.0f}/100**. "
                f"Current status: {readiness_status}."
            )

        else:
            st.info(
                "Betty is still learning. Try asking about focus, today, tasks, birthdays, shopping, budget, household, Bailey or readiness."
            )

# ==================================================
# 14. FAMILY MEMBERS
# ==================================================

with tab7:

    st.header("👨‍👩‍👧 Family Members")

    st.write("Add, update and manage family members, birthdays and important family dates.")

    if not family.empty:

        family_display = family.copy()
        family_display["Birthday"] = pd.to_datetime(family_display["Birthday"], errors="coerce")

        ages = []
        days_until = []
        next_ages = []

        for _, row in family_display.iterrows():

            birthday = row["Birthday"]

            if pd.isna(birthday):
                ages.append("Unknown")
                days_until.append("Unknown")
                next_ages.append("Unknown")
            else:
                birthday_date = birthday.date()

                age = today_date.year - birthday_date.year

                if (today_date.month, today_date.day) < (birthday_date.month, birthday_date.day):
                    age -= 1

                next_birthday = date(today_date.year, birthday_date.month, birthday_date.day)

                if next_birthday < today_date:
                    next_birthday = date(today_date.year + 1, birthday_date.month, birthday_date.day)

                days = (next_birthday - today_date).days

                ages.append(age)
                days_until.append(days)
                next_ages.append(age + 1)

        family_display["Age"] = ages
        family_display["Turning"] = next_ages
        family_display["Days Until Birthday"] = days_until
        family_display["Birthday"] = family_display["Birthday"].dt.strftime("%Y-%m-%d")

        st.subheader("📋 Family List")
        st.dataframe(family_display, use_container_width=True)

        st.divider()

        st.subheader("🎂 Betty Birthday Reminders")

        upcoming_birthdays = family_display[
            family_display["Days Until Birthday"].apply(lambda x: isinstance(x, int) and x <= 30)
        ]

        if not upcoming_birthdays.empty:
            for _, row in upcoming_birthdays.iterrows():
                st.warning(
                    f"🎂 {row['Name']}'s birthday is in {row['Days Until Birthday']} day(s). "
                    f"They will be turning {row['Turning']}."
                )
        else:
            st.success("✅ No birthdays due in the next 30 days.")

    else:
        st.info("No family members added yet.")

    st.divider()

    st.subheader("➕ Add Family Member")

    with st.form("add_family_member"):

        member_name = st.text_input("Name")
        relationship = st.text_input("Relationship", value="Family")

        birthday = st.date_input(
            "Birthday",
            value=date(1990, 1, 1),
            min_value=date(1900, 1, 1),
            max_value=date.today(),
            key="add_member_birthday"
        )

        submitted = st.form_submit_button("Add Family Member")

        if submitted:

            new_member = pd.DataFrame([{
                "Name": member_name,
                "Relationship": relationship,
                "Birthday": str(birthday)
            }])

            family = pd.concat([family, new_member], ignore_index=True)
            save_data(family, FAMILY_FILE)

            st.success("Family member added.")
            st.rerun()

    st.divider()

    st.subheader("✏️ Update Existing Family Member")

    if not family.empty:

        member_to_update = st.selectbox(
            "Select family member to update",
            family["Name"].tolist(),
            key="update_member_select"
        )

        selected_member = family[family["Name"] == member_to_update].iloc[0]

        existing_birthday = pd.to_datetime(
            selected_member["Birthday"],
            errors="coerce"
        )

        if pd.isna(existing_birthday):
            existing_birthday_value = date(1990, 1, 1)
        else:
            existing_birthday_value = existing_birthday.date()

        with st.form("update_family_member"):

            updated_name = st.text_input("Updated Name", value=selected_member["Name"])

            updated_relationship = st.text_input(
                "Updated Relationship",
                value=selected_member["Relationship"]
            )

            updated_birthday = st.date_input(
                "Updated Birthday",
                value=existing_birthday_value,
                min_value=date(1900, 1, 1),
                max_value=date.today(),
                key="update_member_birthday"
            )

            update_submitted = st.form_submit_button("Update Family Member")

            if update_submitted:

                family.loc[
                    family["Name"] == member_to_update,
                    ["Name", "Relationship", "Birthday"]
                ] = [
                    updated_name,
                    updated_relationship,
                    str(updated_birthday)
                ]

                save_data(family, FAMILY_FILE)

                st.success(f"{updated_name} updated successfully.")
                st.rerun()

    st.divider()

    st.subheader("🗑️ Delete Family Member")

    if not family.empty:

        member_to_delete = st.selectbox(
            "Select family member to delete",
            family["Name"].tolist(),
            key="delete_member_select"
        )

        if st.button("Delete Selected Family Member"):

            family = family[family["Name"] != member_to_delete]
            save_data(family, FAMILY_FILE)

            st.success(f"{member_to_delete} removed from FamilyOS.")
            st.rerun()

# ==================================================
# 15. SHOPPING CENTRE
# ==================================================

with tab8:

    st.header("🛒 Shopping Centre")

    st.write(
        "Manage the family shopping list and let Betty suggest meals."
    )

    # --------------------------------------------------
    # Shopping KPIs
    # --------------------------------------------------

    if not shopping.empty:

        open_shopping = shopping[
            shopping["Status"] == "Open"
        ]

        high_priority = open_shopping[
            open_shopping["Priority"] == "High"
        ]

        scol1, scol2, scol3 = st.columns(3)

        with scol1:
            st.metric(
                "Open Items",
                len(open_shopping)
            )

        with scol2:
            st.metric(
                "High Priority",
                len(high_priority)
            )

        with scol3:
            st.metric(
                "Total Items",
                len(shopping)
            )

    st.divider()

    # --------------------------------------------------
    # Shopping List
    # --------------------------------------------------

    st.subheader("🛒 Shopping List")

    st.dataframe(
        shopping,
        use_container_width=True
    )

    st.divider()

    # --------------------------------------------------
    # Add Shopping Item
    # --------------------------------------------------

    st.subheader("➕ Add Shopping Item")

    with st.form("add_shopping_item"):

        item = st.text_input("Item")

        category = st.selectbox(
            "Category",
            [
                "Groceries",
                "Household",
                "Pets",
                "Other"
            ]
        )

        priority = st.selectbox(
            "Priority",
            [
                "Low",
                "Medium",
                "High"
            ]
        )

        added_by = st.text_input(
            "Added By",
            value="Family"
        )

        submitted = st.form_submit_button(
            "Add Item"
        )

        if submitted:

            new_item = pd.DataFrame([{
                "Item": item,
                "Category": category,
                "Priority": priority,
                "Added By": added_by,
                "Status": "Open"
            }])

            shopping = pd.concat(
                [shopping, new_item],
                ignore_index=True
            )

            save_data(
                shopping,
                SHOPPING_FILE
            )

            st.success(
                "Shopping item added."
            )

            st.rerun()

    st.divider()

    # --------------------------------------------------
    # Meal Ideas Centre
    # --------------------------------------------------

    st.subheader("🍽️ Meal Ideas Centre")

    st.dataframe(
        meals,
        use_container_width=True
    )

    st.divider()

    st.subheader("🤖 Betty Dinner Suggestion")

    dinner_meals = meals[
        meals["Meal Type"] == "Dinner"
    ]

    if not dinner_meals.empty:

        suggested_meal = dinner_meals.sample(1).iloc[0]

        st.success(
            f"""
🍽️ Tonight's Suggestion: {suggested_meal['Meal']}

👤 Suggested By: {suggested_meal['Suggested By']}

⭐ Difficulty: {suggested_meal['Difficulty']}

🛒 Ingredients:
{suggested_meal['Ingredients']}
"""
        )

        if st.button(
            "🛒 Add Ingredients To Shopping List",
            key="meal_button"
        ):

            ingredients = [
                x.strip()
                for x in str(
                    suggested_meal["Ingredients"]
                ).split(",")
            ]

            for ingredient in ingredients:

                new_item = pd.DataFrame([{
                    "Item": ingredient,
                    "Category": "Groceries",
                    "Priority": "Medium",
                    "Added By": "Betty",
                    "Status": "Open"
                }])

                shopping = pd.concat(
                    [shopping, new_item],
                    ignore_index=True
                )

            save_data(
                shopping,
                SHOPPING_FILE
            )

            st.success(
                "Ingredients added to shopping list."
            )

            st.rerun()

    st.divider()

    st.subheader("➕ Add Meal Idea")

    with st.form("add_meal_idea"):

        meal_name = st.text_input(
            "Meal Name"
        )

        meal_type = st.selectbox(
            "Meal Type",
            [
                "Breakfast",
                "Lunch",
                "Dinner",
                "Snack"
            ]
        )

        ingredients = st.text_area(
            "Ingredients (comma separated)"
        )

        suggested_by = st.text_input(
            "Suggested By",
            value="Family"
        )

        difficulty = st.selectbox(
            "Difficulty",
            [
                "Easy",
                "Medium",
                "Hard"
            ]
        )

        submitted = st.form_submit_button(
            "Add Meal Idea"
        )

        if submitted:

            new_meal = pd.DataFrame([{
                "Meal": meal_name,
                "Meal Type": meal_type,
                "Ingredients": ingredients,
                "Suggested By": suggested_by,
                "Difficulty": difficulty
            }])

            meals = pd.concat(
                [meals, new_meal],
                ignore_index=True
            )

            save_data(
                meals,
                MEALS_FILE
            )

            st.success(
                "Meal idea added."
            )

            st.rerun()

# ==================================================
# 16. BUDGET CENTRE
# ==================================================

with tab9:

    st.header("💰 Budget Centre")

    st.write("Track simple family spending and let Betty highlight budget pressure.")

    bcol1, bcol2, bcol3, bcol4 = st.columns(4)

    with bcol1:
        st.metric("Monthly Budget", f"£{monthly_budget_total:,.2f}")

    with bcol2:
        st.metric("Actual Spend", f"£{monthly_spend_total:,.2f}")

    with bcol3:
        st.metric("Remaining", f"£{remaining_budget:,.2f}")

    with bcol4:
        st.metric("Budget Used", f"{budget_used_pct:.1f}%")

    st.divider()

    st.subheader("📊 Budget Categories")

    spend_by_category = budget.groupby("Category")["Amount (£)"].sum().reset_index() if not budget.empty else pd.DataFrame(columns=["Category", "Amount (£)"])

    budget_summary = budget_categories.merge(
        spend_by_category,
        on="Category",
        how="left"
    )

    budget_summary["Amount (£)"] = budget_summary["Amount (£)"].fillna(0)
    budget_summary["Remaining (£)"] = budget_summary["Monthly Budget (£)"] - budget_summary["Amount (£)"]
    budget_summary["Used %"] = (
        budget_summary["Amount (£)"] / budget_summary["Monthly Budget (£)"] * 100
    ).round(1)

    st.dataframe(
        budget_summary,
        use_container_width=True
    )

    st.divider()

    st.subheader("📋 Expense Tracker")

    st.dataframe(
        budget,
        use_container_width=True
    )

    st.divider()

    st.subheader("🤖 Betty Budget Insights")

    if budget_used_pct >= 100:
        st.error("🔴 Monthly spending is over budget. Betty recommends reviewing discretionary spending.")
    elif budget_used_pct >= 80:
        st.warning("🟠 Monthly spending is above 80% of budget. Keep an eye on remaining spend.")
    else:
        st.success("🟢 Family spending is currently within budget.")

    over_budget_categories = budget_summary[budget_summary["Remaining (£)"] < 0]

    if not over_budget_categories.empty:
        for _, row in over_budget_categories.iterrows():
            st.warning(
                f"⚠️ {row['Category']} is over budget by £{abs(row['Remaining (£)']):,.2f}."
            )

    highest_spend = budget_summary.sort_values("Amount (£)", ascending=False).iloc[0]

    st.info(
        f"💡 Highest spend category is **{highest_spend['Category']}** "
        f"at £{highest_spend['Amount (£)']:,.2f}."
    )

    st.divider()

    st.subheader("➕ Add Expense")

    with st.form("add_expense"):

        expense_date = st.date_input("Expense Date", value=date.today())
        expense_category = st.selectbox(
            "Category",
            budget_categories["Category"].tolist()
        )
        expense_description = st.text_input("Description")
        expense_amount = st.number_input(
            "Amount (£)",
            min_value=0.0,
            value=10.0,
            step=1.0
        )

        submitted = st.form_submit_button("Add Expense")

        if submitted:

            new_expense = pd.DataFrame([{
                "Date": str(expense_date),
                "Category": expense_category,
                "Description": expense_description,
                "Amount (£)": expense_amount
            }])

            budget = pd.concat([budget, new_expense], ignore_index=True)
            save_data(budget, BUDGET_FILE)

            st.success("Expense added.")
            st.rerun()


# ==================================================
# 17. PET CENTRE
# ==================================================

with tab10:

    st.header("🐶 Pet Centre")

    st.write("Track pet care, vet appointments, insurance, vaccinations, medication and weight.")

    if not pets.empty:

        pets_display = pets.copy()

        pets_display["Birthday"] = pd.to_datetime(pets_display["Birthday"], errors="coerce")
        pets_display["Insurance Renewal"] = pd.to_datetime(pets_display["Insurance Renewal"], errors="coerce")
        pets_display["Next Vaccination"] = pd.to_datetime(pets_display["Next Vaccination"], errors="coerce")

        insurance_days = []
        vaccination_days = []

        for _, row in pets_display.iterrows():

            insurance_date = row["Insurance Renewal"]
            vaccination_date = row["Next Vaccination"]

            insurance_days.append((insurance_date.date() - date.today()).days if not pd.isna(insurance_date) else "Unknown")
            vaccination_days.append((vaccination_date.date() - date.today()).days if not pd.isna(vaccination_date) else "Unknown")

        pets_display["Days Until Insurance Renewal"] = insurance_days
        pets_display["Days Until Vaccination"] = vaccination_days

        display_pets = pets_display.copy()
        display_pets["Birthday"] = display_pets["Birthday"].dt.strftime("%Y-%m-%d")
        display_pets["Insurance Renewal"] = display_pets["Insurance Renewal"].dt.strftime("%Y-%m-%d")
        display_pets["Next Vaccination"] = display_pets["Next Vaccination"].dt.strftime("%Y-%m-%d")

        st.subheader("📋 Pet Profiles")

        st.dataframe(
            display_pets,
            use_container_width=True
        )

        st.divider()

        st.subheader("🤖 Betty Pet Insights")

        reminders_found = False

        for _, row in pets_display.iterrows():

            pet_name = row["Pet Name"]

            if isinstance(row["Days Until Insurance Renewal"], int) and row["Days Until Insurance Renewal"] <= 30:
                st.warning(
                    f"🛡️ {pet_name}'s insurance renewal is due in {row['Days Until Insurance Renewal']} day(s)."
                )
                reminders_found = True

            if isinstance(row["Days Until Vaccination"], int) and row["Days Until Vaccination"] <= 30:
                st.warning(
                    f"💉 {pet_name}'s vaccination is due in {row['Days Until Vaccination']} day(s)."
                )
                reminders_found = True

        if not reminders_found:
            st.success("✅ No urgent pet reminders in the next 30 days.")

    else:
        st.info("No pets added yet.")

    st.divider()

    st.subheader("➕ Add Pet")

    with st.form("add_pet"):

        pet_name = st.text_input("Pet Name")
        animal = st.text_input("Animal", value="Dog")
        breed = st.text_input("Breed", value="Unknown")

        birthday = st.date_input(
            "Birthday",
            value=date(2020, 1, 1),
            min_value=date(1990, 1, 1),
            max_value=date.today(),
            key="add_pet_birthday"
        )

        vet = st.text_input("Vet", value="Local Vet")

        insurance_renewal = st.date_input(
            "Insurance Renewal",
            value=date.today(),
            key="add_pet_insurance"
        )

        next_vaccination = st.date_input(
            "Next Vaccination",
            value=date.today(),
            key="add_pet_vaccination"
        )

        medication = st.text_input("Medication", value="None")

        weight = st.number_input(
            "Weight (kg)",
            min_value=0.0,
            value=0.0,
            step=0.1,
            key="add_pet_weight"
        )

        notes = st.text_area("Notes")

        submitted = st.form_submit_button("Add Pet")

        if submitted:

            new_pet = pd.DataFrame([{
                "Pet Name": pet_name,
                "Animal": animal,
                "Breed": breed,
                "Birthday": str(birthday),
                "Vet": vet,
                "Insurance Renewal": str(insurance_renewal),
                "Next Vaccination": str(next_vaccination),
                "Medication": medication,
                "Weight (kg)": weight,
                "Notes": notes
            }])

            pets = pd.concat([pets, new_pet], ignore_index=True)
            save_data(pets, PETS_FILE)

            st.success("Pet added.")
            st.rerun()

    st.divider()

    st.subheader("✏️ Update Pet")

    if not pets.empty:

        pet_to_update = st.selectbox(
            "Select pet to update",
            pets["Pet Name"].tolist(),
            key="update_pet_select"
        )

        selected_pet = pets[pets["Pet Name"] == pet_to_update].iloc[0]

        existing_birthday = pd.to_datetime(selected_pet["Birthday"], errors="coerce")
        existing_insurance = pd.to_datetime(selected_pet["Insurance Renewal"], errors="coerce")
        existing_vaccination = pd.to_datetime(selected_pet["Next Vaccination"], errors="coerce")

        birthday_value = existing_birthday.date() if not pd.isna(existing_birthday) else date(2020, 1, 1)
        insurance_value = existing_insurance.date() if not pd.isna(existing_insurance) else date.today()
        vaccination_value = existing_vaccination.date() if not pd.isna(existing_vaccination) else date.today()

        with st.form("pets_update_form"):

            updated_name = st.text_input("Updated Pet Name", value=selected_pet["Pet Name"])
            updated_animal = st.text_input("Updated Animal", value=selected_pet["Animal"])
            updated_breed = st.text_input("Updated Breed", value=selected_pet["Breed"])

            updated_birthday = st.date_input(
                "Updated Birthday",
                value=birthday_value,
                min_value=date(1990, 1, 1),
                max_value=date.today(),
                key="update_pet_birthday"
            )

            updated_vet = st.text_input("Updated Vet", value=selected_pet["Vet"])

            updated_insurance = st.date_input(
                "Updated Insurance Renewal",
                value=insurance_value,
                key="update_pet_insurance"
            )

            updated_vaccination = st.date_input(
                "Updated Next Vaccination",
                value=vaccination_value,
                key="update_pet_vaccination"
            )

            updated_medication = st.text_input("Updated Medication", value=selected_pet["Medication"])

            updated_weight = st.number_input(
                "Updated Weight (kg)",
                min_value=0.0,
                value=float(selected_pet["Weight (kg)"]),
                step=0.1,
                key="update_pet_weight"
            )

            updated_notes = st.text_area("Updated Notes", value=str(selected_pet["Notes"]))

            update_submitted = st.form_submit_button("Update Pet")

            if update_submitted:

                pets.loc[
                    pets["Pet Name"] == pet_to_update,
                    [
                        "Pet Name",
                        "Animal",
                        "Breed",
                        "Birthday",
                        "Vet",
                        "Insurance Renewal",
                        "Next Vaccination",
                        "Medication",
                        "Weight (kg)",
                        "Notes"
                    ]
                ] = [
                    updated_name,
                    updated_animal,
                    updated_breed,
                    str(updated_birthday),
                    updated_vet,
                    str(updated_insurance),
                    str(updated_vaccination),
                    updated_medication,
                    updated_weight,
                    updated_notes
                ]

                save_data(pets, PETS_FILE)

                st.success(f"{updated_name} updated successfully.")
                st.rerun()

    else:
        st.info("No pets available to update.")

    st.divider()

    st.subheader("🗑️ Delete Pet")

    if not pets.empty:

        pet_to_delete = st.selectbox(
            "Select pet to delete",
            pets["Pet Name"].tolist(),
            key="delete_pet_select"
        )

        if st.button("Delete Selected Pet"):

            pets = pets[pets["Pet Name"] != pet_to_delete]
            save_data(pets, PETS_FILE)

            st.success(f"{pet_to_delete} removed from Pet Centre.")
            st.rerun()

    else:
        st.info("No pets available to delete.")

# ==================================================
# 18. HOUSEHOLD MAINTENANCE CENTRE
# ==================================================

with tab11:

    st.header("🏡 Household Maintenance Centre")

    st.write("Track household renewals, services, vehicle dates and important maintenance reminders.")

    if not household.empty:

        household_display = household.copy()
        household_display["Due Date"] = pd.to_datetime(
            household_display["Due Date"],
            errors="coerce"
        )

        days_until_due = []

        for _, row in household_display.iterrows():

            due_date = row["Due Date"]

            if pd.isna(due_date):
                days_until_due.append("Unknown")
            else:
                days_until_due.append(
                    (due_date.date() - date.today()).days
                )

        household_display["Days Until Due"] = days_until_due

        display_household = household_display.copy()
        display_household["Due Date"] = display_household["Due Date"].dt.strftime("%Y-%m-%d")

        st.subheader("📋 Household Items")

        st.dataframe(
            display_household,
            use_container_width=True
        )

        st.divider()

        st.subheader("🤖 Betty Household Insights")

        reminders_found = False

        for _, row in household_display.iterrows():

            item = row["Item"]
            days = row["Days Until Due"]

            if isinstance(days, int) and days < 0:
                st.error(
                    f"🔴 {item} is overdue by {abs(days)} day(s)."
                )
                reminders_found = True

            elif isinstance(days, int) and days <= 30:
                st.warning(
                    f"🟠 {item} is due in {days} day(s)."
                )
                reminders_found = True

        if not reminders_found:
            st.success("✅ No household items due in the next 30 days.")

    else:
        st.info("No household items added yet.")

    st.divider()

    st.subheader("➕ Add Household Item")

    with st.form("add_household_item"):

        item = st.text_input("Item")
        category = st.selectbox(
            "Category",
            ["Home", "Vehicle", "Insurance", "Utilities", "Finance", "Other"]
        )

        due_date = st.date_input(
            "Due Date",
            value=date.today(),
            key="add_household_due_date"
        )

        owner = st.text_input("Owner", value="Nick")

        status = st.selectbox(
            "Status",
            ["Open", "Complete"]
        )

        notes = st.text_area("Notes")

        submitted = st.form_submit_button("Add Household Item")

        if submitted:

            new_item = pd.DataFrame([{
                "Item": item,
                "Category": category,
                "Due Date": str(due_date),
                "Owner": owner,
                "Status": status,
                "Notes": notes
            }])

            household = pd.concat([household, new_item], ignore_index=True)
            save_data(household, HOUSEHOLD_FILE)

            st.success("Household item added.")
            st.rerun()

# ==================================================
# 19. NOTIFICATION CENTRE
# ==================================================

with tab12:

    st.header("🔔 Betty Notification Centre")

    st.write(
        "Everything that needs your attention in one place."
    )

    notifications_found = False

    # ==================================================
    # TASKS
    # ==================================================

    st.subheader("✅ Tasks")

    if len(open_tasks) > 0:

        notifications_found = True

        st.warning(
            f"There are currently {len(open_tasks)} open task(s)."
        )

        st.dataframe(
            open_tasks,
            use_container_width=True
        )

    else:
        st.success("No open tasks.")

    st.divider()

    # ==================================================
    # BIRTHDAYS
    # ==================================================

    st.subheader("🎂 Upcoming Birthdays")

    if not birthday_reminders.empty:

        notifications_found = True

        birthday_display = birthday_reminders[
            ["Name", "Days Until Birthday"]
        ]

        st.dataframe(
            birthday_display,
            use_container_width=True
        )

    else:
        st.success("No birthdays due soon.")

    st.divider()

    # ==================================================
    # HOUSEHOLD
    # ==================================================

    st.subheader("🏡 Household Reminders")

    if not household_reminders.empty:

        notifications_found = True

        household_display = household_reminders[
            ["Item", "Due Date", "Days Until Due"]
        ]

        st.dataframe(
            household_display,
            use_container_width=True
        )

    else:
        st.success("No household reminders.")

    st.divider()

    # ==================================================
    # PETS
    # ==================================================

    st.subheader("🐶 Pet Reminders")

    pet_notifications = []

    if not pets.empty:

        pets_check = pets.copy()

        pets_check["Insurance Renewal"] = pd.to_datetime(
            pets_check["Insurance Renewal"],
            errors="coerce"
        )

        pets_check["Next Vaccination"] = pd.to_datetime(
            pets_check["Next Vaccination"],
            errors="coerce"
        )

        for _, row in pets_check.iterrows():

            pet_name = row["Pet Name"]

            if not pd.isna(row["Insurance Renewal"]):

                insurance_days = (
                    row["Insurance Renewal"].date() - date.today()
                ).days

                if insurance_days <= 30:

                    pet_notifications.append({
                        "Pet": pet_name,
                        "Reminder": "Insurance Renewal",
                        "Days": insurance_days
                    })

            if not pd.isna(row["Next Vaccination"]):

                vaccination_days = (
                    row["Next Vaccination"].date() - date.today()
                ).days

                if vaccination_days <= 30:

                    pet_notifications.append({
                        "Pet": pet_name,
                        "Reminder": "Vaccination",
                        "Days": vaccination_days
                    })

    if len(pet_notifications) > 0:

        notifications_found = True

        st.dataframe(
            pd.DataFrame(pet_notifications),
            use_container_width=True
        )

    else:
        st.success("No pet reminders.")

    st.divider()

    # ==================================================
    # BUDGET
    # ==================================================

    st.subheader("💰 Budget Alerts")

    if budget_used_pct >= 80:

        notifications_found = True

        st.warning(
            f"Budget is currently {budget_used_pct:.1f}% used."
        )

    else:
        st.success(
            f"Budget is healthy at {budget_used_pct:.1f}% used."
        )

    st.divider()

    # ==================================================
    # BETTY SUMMARY
    # ==================================================

    st.subheader("🤖 Betty Summary")

    if notifications_found:

        st.info(
            """
Betty has identified items that need attention.

Review the notifications above and tackle the highest priority items first.
"""
        )

    else:

        st.success(
            """
🎉 Everything looks under control.

No major notifications detected.
"""
        )

# ==================================================
# 20. BACKUP CENTRE
# ==================================================

with tab13:

    st.header("💾 Backup Centre")

    st.write("Export key FamilyOS data so you can keep a safe copy.")

    st.divider()

    st.subheader("📤 Export Data")

    st.download_button(
        label="Download Family Members",
        data=family.to_csv(index=False),
        file_name="family_members_backup.csv",
        mime="text/csv"
    )

    st.download_button(
        label="Download Tasks",
        data=tasks.to_csv(index=False),
        file_name="tasks_backup.csv",
        mime="text/csv"
    )

    st.download_button(
        label="Download Shopping List",
        data=shopping.to_csv(index=False),
        file_name="shopping_backup.csv",
        mime="text/csv"
    )

    st.download_button(
        label="Download Budget",
        data=budget.to_csv(index=False),
        file_name="budget_backup.csv",
        mime="text/csv"
    )

    st.download_button(
        label="Download Pets",
        data=pets.to_csv(index=False),
        file_name="pets_backup.csv",
        mime="text/csv"
    )

    st.download_button(
        label="Download Household",
        data=household.to_csv(index=False),
        file_name="household_backup.csv",
        mime="text/csv"
    )

    st.divider()

    st.info("This is the first Beta backup version. Later we can add full one-click backup and restore.")


