import os
import json
import urllib.request
from io import BytesIO
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
from flask_sqlalchemy import SQLAlchemy
from flask_bcrypt import Bcrypt
from werkzeug.utils import secure_filename
import cloudinary
import cloudinary.uploader
import cloudinary.api
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from datetime import datetime
import pytz
from pptx.oxml.xmlchemy import OxmlElement

def add_highlight(run, color_hex):
    rPr = run._r.get_or_add_rPr()
    highlight = OxmlElement('a:highlight')
    srgbClr = OxmlElement('a:srgbClr')
    srgbClr.set('val', color_hex)
    highlight.append(srgbClr)
    rPr.append(highlight)

MENTOR_NAMES = {
    'vijay@123': 'Vijayavaran',
    'mentor2@123': 'Mr. Mentor 2',
    'Ashwin@123': 'Mr. V. Ashwin',
    'admin@123': 'Administrator'
}

app = Flask(__name__)
app.config['SECRET_KEY'] = 'cinematic_secret_key_123'
uri = os.environ.get("DATABASE_URL", "sqlite:///database.db")
if uri.startswith("postgres://"):
    uri = uri.replace("postgres://", "postgresql://", 1)
app.config['SQLALCHEMY_DATABASE_URI'] = uri

# Use QueuePool for better concurrency. NullPool opens a new connection per query,
# which causes connection exhaustion / 500 errors when many students use the app at once.
if 'postgresql' in uri or 'postgres' in uri:
    app.config['SQLALCHEMY_ENGINE_OPTIONS'] = {
        'pool_size': 10,
        'max_overflow': 20,
        'pool_timeout': 30,
        'pool_recycle': 280,
        'pool_pre_ping': True,
        'connect_args': {
            'sslmode': 'require',
            'connect_timeout': 10,
        },
    }
else:
    app.config['SQLALCHEMY_ENGINE_OPTIONS'] = {
        'pool_pre_ping': True,
        'pool_recycle': 300,
    }
app.config['UPLOAD_FOLDER'] = 'static/uploads'
app.config['ALLOWED_EXTENSIONS'] = {'png', 'jpg', 'jpeg', 'gif'}
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

db = SQLAlchemy(app)
bcrypt = Bcrypt(app)

# Models
class User(db.Model):
    username = db.Column(db.String(50), primary_key=True)
    password_hash = db.Column(db.String(100), nullable=False)
    role = db.Column(db.String(20), nullable=False) # 'faculty' or 'student'

class StudentDetail(db.Model):
    reg_num = db.Column(db.String(50), primary_key=True)
    name = db.Column(db.String(100))
    course = db.Column(db.String(100))
    mentor_username = db.Column(db.String(50), db.ForeignKey('user.username'))
    # JSON strings for dynamic data
    slot_info = db.Column(db.Text, default='[]') # List of slot names
    attendance_data = db.Column(db.Text, default='{}') # {"Slot A": 81}
    marks_data = db.Column(db.Text, default='{}') # {"Slot A": {"model": 20}}
    
    registered_new_course = db.Column(db.String(200))
    online_course = db.Column(db.String(200))
    event_participation = db.Column(db.Text)
    additional_description = db.Column(db.Text)
    results_data = db.Column(db.Text, default='[]')  # [{"subject": "Maths", "grade": "A"}]
    custom_advisory = db.Column(db.Text)
    photo_path = db.Column(db.String(200))
    last_updated = db.Column(db.DateTime, nullable=True)

    def get_attendance(self):
        try: return json.loads(self.attendance_data)
        except: return {}
    
    def get_marks(self):
        try: return json.loads(self.marks_data)
        except: return {}
    
    def get_slots(self):
        try: return json.loads(self.slot_info)
        except: return []

class GlobalAdvisory(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    content = db.Column(db.Text, nullable=False, default="All students are advised to pay their 2nd-year tuition fees on time through the Viana Portal.\nAdditionally, kindly upload your recent passport-size photograph to your Viana profile at the earliest, if you have not already done so....")

class GlobalMentorObservation(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    content = db.Column(db.Text, nullable=False, default="I personally advised the student to concentrate more on study and skill development. The student is currently attending an online course to improve their technical skills, which is really appreciable.")

class GlobalSettings(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    show_grades_in_ppt = db.Column(db.Boolean, nullable=False, default=False)

# Helper functions
def delete_photo(photo_path):
    if not photo_path:
        return
    if 'cloudinary.com' in photo_path:
        try:
            public_id = "simats_profiles/" + photo_path.split('/')[-1].split('.')[0]
            cloudinary.uploader.destroy(public_id)
        except Exception as e:
            print(f'Failed to delete from Cloudinary: {e}')
    elif os.path.exists(photo_path):
        try:
            os.unlink(photo_path)
        except Exception as e:
            print(f'Failed to delete local photo: {e}')

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def seed_db():
    # Faculty: vijay@123
    if not User.query.get('vijay@123'):
        old_user = User.query.get('mentor1@123')
        if old_user:
            new_user = User(username='vijay@123', password_hash=old_user.password_hash, role='faculty')
            db.session.add(new_user)
            db.session.flush()
            StudentDetail.query.filter_by(mentor_username='mentor1@123').update({'mentor_username': 'vijay@123'})
            db.session.flush()
            db.session.delete(old_user)
        else:
            db.session.add(User(
                username='vijay@123',
                password_hash=bcrypt.generate_password_hash('welcome').decode('utf-8'),
                role='faculty'
            ))

    # Faculty: mentor2@123
    if not User.query.get('mentor2@123'):
        db.session.add(User(
            username='mentor2@123',
            password_hash=bcrypt.generate_password_hash('welcome').decode('utf-8'),
            role='faculty'
        ))

    # Faculty: Ashwin@123
    if not User.query.get('Ashwin@123'):
        db.session.add(User(
            username='Ashwin@123',
            password_hash=bcrypt.generate_password_hash('welcome').decode('utf-8'),
            role='faculty'
        ))

    # Faculty: admin@123
    if not User.query.get('admin@123'):
        db.session.add(User(
            username='admin@123',
            password_hash=bcrypt.generate_password_hash('welcome').decode('utf-8'),
            role='faculty'
        ))

    # Demo student: 24REG01
    if not User.query.get('24REG01'):
        db.session.add(User(
            username='24REG01',
            password_hash=bcrypt.generate_password_hash('student123').decode('utf-8'),
            role='student'
        ))
        db.session.add(StudentDetail(
            reg_num='24REG01',
            name='Ibrahim',
            course='CSA0708 - Computer Networks',
            mentor_username='vijay@123',
            slot_info=json.dumps(['Slot A', 'Slot B']),
            attendance_data=json.dumps({'Slot A': 81, 'Slot B': 98}),
            marks_data=json.dumps({'Slot A': {'model': '20', 'test1': '20', 'avg': '15'}}),
            last_updated=datetime.now(pytz.timezone('Asia/Kolkata')).replace(tzinfo=None)
        ))

    # Demo student: Ibrahim@123
    if not User.query.get('Ibrahim@123'):
        db.session.add(User(
            username='Ibrahim@123',
            password_hash=bcrypt.generate_password_hash('welcome').decode('utf-8'),
            role='student'
        ))

    db.session.commit()

with app.app_context():
    db.create_all()  # Only creates tables if they don't exist — never wipes data
    try:
        from sqlalchemy import inspect
        inspector = inspect(db.engine)
        columns = [col['name'] for col in inspector.get_columns('student_detail')]
        if 'mentor_username' not in columns:
            db.session.execute(db.text('ALTER TABLE student_detail ADD COLUMN mentor_username VARCHAR(50)'))
            db.session.commit()
        if 'last_updated' not in columns:
            col_type = 'TIMESTAMP' if db.engine.name == 'postgresql' else 'DATETIME'
            db.session.execute(db.text(f'ALTER TABLE student_detail ADD COLUMN last_updated {col_type}'))
            db.session.commit()
        if 'results_data' not in columns:
            db.session.execute(db.text('ALTER TABLE student_detail ADD COLUMN results_data TEXT DEFAULT \'[]\''))
            db.session.commit()
    except Exception as e:
        db.session.rollback()
        print(f"Migration error: {e}")
    seed_db()

# Routes
@app.route('/')
def index():
    if 'user' in session:
        if session['role'] == 'faculty':
            return redirect(url_for('faculty_dashboard'))
        return redirect(url_for('student_dashboard'))
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        user = User.query.filter_by(username=username).first()
        
        if user and bcrypt.check_password_hash(user.password_hash, password):
            session['user'] = user.username
            session['role'] = user.role
            if user.role == 'faculty':
                return redirect(url_for('faculty_dashboard'))
            return redirect(url_for('student_dashboard'))
        flash('Invalid username or password', 'danger')
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/change_password', methods=['POST'])
def change_password():
    if 'user' not in session:
        return redirect(url_for('login'))
    
    current_password = request.form['current_password']
    new_password = request.form['new_password']
    
    user = User.query.get(session['user'])
    if user and bcrypt.check_password_hash(user.password_hash, current_password):
        user.password_hash = bcrypt.generate_password_hash(new_password).decode('utf-8')
        db.session.commit()
        flash('Password changed successfully!', 'success')
    else:
        flash('Incorrect current password', 'danger')
    
    return redirect(url_for('faculty_dashboard' if session['role'] == 'faculty' else 'student_dashboard'))

@app.route('/student', methods=['GET', 'POST'])
def student_dashboard():
    if 'user' not in session or session['role'] != 'student':
        return redirect(url_for('login'))
    
    student = StudentDetail.query.get(session['user'])
    if not student:
        student = StudentDetail(reg_num=session['user'], name=session['user'])
        db.session.add(student)
        db.session.commit()

    if request.method == 'POST':
        try:
            student.name = (request.form.get('name') or '').strip() or student.name
            student.course = (request.form.get('course') or '').strip() or student.course

            # Dynamic Slots handling
            slots = [s.strip() for s in request.form.getlist('slot_names[]') if s and s.strip()]
            if not slots:
                slots = student.get_slots() or []
            student.slot_info = json.dumps(slots)

            att_data = {}
            marks_data = {}
            current_attendance = student.get_attendance()
            current_marks = student.get_marks()
            for slot in slots:
                att_val = request.form.get(f'att_{slot}', current_attendance.get(slot, 0))
                att_data[slot] = str(att_val).strip() if att_val is not None else 0
                marks_data[slot] = {
                    'model': (request.form.get(f'model_{slot}') or current_marks.get(slot, {}).get('model', '-')).strip(),
                    'test1': (request.form.get(f'test1_{slot}') or current_marks.get(slot, {}).get('test1', '-')).strip(),
                    'test2': (request.form.get(f'test2_{slot}') or current_marks.get(slot, {}).get('test2', '-')).strip(),
                    'avg': (request.form.get(f'avg_{slot}') or current_marks.get(slot, {}).get('avg', '-')).strip(),
                    'total_marks': (request.form.get(f'total_marks_{slot}') or current_marks.get(slot, {}).get('total_marks', '')).strip(),
                    'course': (request.form.get(f'course_{slot}') or current_marks.get(slot, {}).get('course', '')).strip()
                }

            student.attendance_data = json.dumps(att_data)
            student.marks_data = json.dumps(marks_data)

            student.registered_new_course = (request.form.get('registered_new_course') or '').strip()
            student.online_course = (request.form.get('online_course') or '').strip()
            student.event_participation = (request.form.get('event_participation') or '').strip()
            student.additional_description = (request.form.get('description') or '').strip()

            # Results data — save subject/grade entries from My Results tab
            result_subjects = request.form.getlist('result_subject[]')
            result_grades = request.form.getlist('result_grade[]')
            if result_subjects:  # Only update if results fields were submitted
                results = [{'subject': s.strip(), 'grade': g.strip()} for s, g in zip(result_subjects, result_grades) if s and s.strip()]
                student.results_data = json.dumps(results)

            file = request.files.get('photo')
            if file and file.filename and allowed_file(file.filename):
                if os.environ.get('CLOUDINARY_URL'):
                    try:
                        upload_result = cloudinary.uploader.upload(file, folder="simats_profiles")
                        student.photo_path = upload_result.get('secure_url')
                    except Exception as e:
                        flash(f'Failed to upload photo to cloud storage: {str(e)}', 'warning')
                else:
                    try:
                        filename = secure_filename(f"{session['user']}_{file.filename}")
                        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                        file.save(filepath)
                        student.photo_path = filepath
                    except Exception as e:
                        print(f"Local photo upload warning: {e}")

            # Safe Kolkata timestamp calculation
            from datetime import timedelta
            student.last_updated = datetime.utcnow() + timedelta(hours=5, minutes=30)
            
            db.session.commit()
            flash('Details updated successfully!', 'success')
        except Exception as e:
            db.session.rollback()
            print(f"Student submit error: {e}")
            flash('An error occurred while saving your details. Please try again.', 'danger')

        return redirect(url_for('student_dashboard'))
    
    # Passing current data as dicts
    try:
        results = json.loads(student.results_data) if student.results_data else []
    except:
        results = []
    return render_template('student.html', student=student,
                           attendance=student.get_attendance(),
                           marks=student.get_marks(),
                           slots=student.get_slots(),
                           results=results)

@app.route('/add_student', methods=['POST'])
def add_student():
    if 'user' not in session or session['role'] != 'faculty':
        return redirect(url_for('login'))
    
    reg_num = request.form['reg_num']
    name = request.form['name']
    password = request.form['password']
    
    if User.query.get(reg_num):
        flash('Student with this Registration Number already exists.', 'danger')
    else:
        new_user = User(
            username=reg_num,
            password_hash=bcrypt.generate_password_hash(password).decode('utf-8'),
            role='student'
        )
        from datetime import timedelta
        new_detail = StudentDetail(
            reg_num=reg_num,
            name=name,
            mentor_username=session['user'],
            last_updated=datetime.utcnow() + timedelta(hours=5, minutes=30)
        )
        db.session.add(new_user)
        db.session.add(new_detail)
        db.session.commit()
        flash(f'Student {name} added successfully!', 'success')
    
    return redirect(url_for('faculty_dashboard'))

@app.route('/edit_student_faculty/<reg_num>', methods=['POST'])
def edit_student_faculty(reg_num):
    if 'user' not in session or session['role'] != 'faculty':
        return redirect(url_for('login'))
    
    student = StudentDetail.query.filter_by(reg_num=reg_num, mentor_username=session['user']).first()
    if not student:
        student = StudentDetail.query.get(reg_num)
        if student and not student.mentor_username:
            student.mentor_username = session['user']

    if student:
        student.name = request.form.get('name', student.name) or student.name
        student.course = request.form.get('course', student.course) or student.course
        student.registered_new_course = request.form.get('registered_new_course', student.registered_new_course or '')
        student.online_course = request.form.get('online_course', student.online_course or '')
        student.event_participation = request.form.get('event_participation', student.event_participation or '')
        student.additional_description = request.form.get('description', student.additional_description or '')
        student.custom_advisory = request.form.get('custom_advisory', student.custom_advisory or '')

        # Save results from faculty edit modal
        result_subjects = request.form.getlist('result_subject[]')
        result_grades = request.form.getlist('result_grade[]')
        results = [{'subject': s.strip(), 'grade': g.strip()} for s, g in zip(result_subjects, result_grades) if s.strip()]
        student.results_data = json.dumps(results)

        from datetime import timedelta
        student.last_updated = datetime.utcnow() + timedelta(hours=5, minutes=30)
        db.session.commit()
        flash(f'Details for {student.name or reg_num} updated successfully!', 'success')
    else:
        flash('Student not found.', 'danger')
        
    return redirect(url_for('faculty_dashboard'))

@app.route('/update_global_advisory', methods=['POST'])
def update_global_advisory():
    if 'user' not in session or session['role'] != 'faculty':
        return redirect(url_for('login'))
    
    advisory_text = request.form.get('global_advisory', '').strip()
    global_adv = GlobalAdvisory.query.first()
    if not global_adv:
        global_adv = GlobalAdvisory(id=1, content=advisory_text)
        db.session.add(global_adv)
    else:
        global_adv.content = advisory_text
    
    db.session.commit()
    flash('Global Institutional Advisory updated for all students!', 'success')
    return redirect(url_for('faculty_dashboard'))

@app.route('/update_global_observation', methods=['POST'])
def update_global_observation():
    if 'user' not in session or session['role'] != 'faculty':
        return redirect(url_for('login'))
    
    observation_text = request.form.get('global_observation', '').strip()
    global_obs = GlobalMentorObservation.query.first()
    if not global_obs:
        global_obs = GlobalMentorObservation(id=1, content=observation_text)
        db.session.add(global_obs)
    else:
        global_obs.content = observation_text
    
    db.session.commit()
    flash('Global Mentor Observation updated for all students!', 'success')
    return redirect(url_for('faculty_dashboard'))

@app.route('/faculty')
def faculty_dashboard():
    if 'user' not in session or session['role'] != 'faculty':
        return redirect(url_for('login'))
    
    # Filter students strictly by logged in mentor
    students = StudentDetail.query.filter_by(mentor_username=session['user']).all()
    global_adv = GlobalAdvisory.query.first()
    global_advisory_text = global_adv.content if global_adv else ""
    global_obs = GlobalMentorObservation.query.first()
    global_observation_text = global_obs.content if global_obs else ""
    
    # Stats filtered per mentor
    total_att_a = 0
    count_a = 0
    for s in students:
        att = s.get_attendance()
        if 'Slot A' in att:
            try:
                total_att_a += int(att['Slot A'] or 0)
                count_a += 1
            except (ValueError, TypeError):
                pass
    
    stats = {
        'total_students': len(students),
        'reports_generated': 152,
        'avg_attendance_a': int(total_att_a / count_a) if count_a > 0 else 0,
        'avg_attendance_b': 92
    }

    students_json = [
        {
            'reg_num': s.reg_num,
            'name': s.name,
            'course': s.course,
            'attendance_data': s.attendance_data,
            'marks_data': s.marks_data,
            'slot_info': s.slot_info,
            'registered_new_course': s.registered_new_course,
            'online_course': s.online_course,
            'event_participation': s.event_participation,
            'additional_description': s.additional_description,
            'custom_advisory': s.custom_advisory,
            'photo_path': s.photo_path,
            'last_updated': s.last_updated.strftime('%d-%b-%Y %I:%M %p') if s.last_updated else 'N/A',
        }
        for s in students
    ]

    mentor_display = MENTOR_NAMES.get(session['user'], session['user'])

    return render_template('faculty.html', students=students, students_json=students_json, stats=stats,
                           mentor_name=mentor_display,
                           global_advisory=global_advisory_text,
                           global_observation=global_observation_text,
                           show_grades_in_ppt=(GlobalSettings.query.first().show_grades_in_ppt if GlobalSettings.query.first() else False))

@app.route('/toggle_grade_ppt', methods=['POST'])
def toggle_grade_ppt():
    if 'user' not in session or session['role'] != 'faculty':
        return redirect(url_for('login'))
    settings = GlobalSettings.query.first()
    if not settings:
        settings = GlobalSettings(id=1, show_grades_in_ppt=True)
        db.session.add(settings)
    else:
        settings.show_grades_in_ppt = not settings.show_grades_in_ppt
    db.session.commit()
    return redirect(url_for('faculty_dashboard'))

@app.route('/generate_report')
def generate_report():
    if 'user' not in session or session['role'] != 'faculty':
        return redirect(url_for('login'))
    
    students = StudentDetail.query.filter_by(mentor_username=session['user']).all()
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    mentor_name = MENTOR_NAMES.get(session['user'], session['user'])

    for student in students:
        slots = student.get_slots()
        att_data = student.get_attendance()
        marks_data = student.get_marks()
        
        # --- SLIDE 1: ACADEMIC PERFORMANCE ---
        slide_layout = prs.slide_layouts[6] # Blank
        slide1 = prs.slides.add_slide(slide_layout)
        
        # Banner Table for Name/Reg/Mentor
        banner_tbl = slide1.shapes.add_table(1, 6, Inches(0.2), Inches(1.4), Inches(13.0), Inches(0.5)).table
        banner_content = [
            "Mentee Name", student.name or 'N/A',
            "Reg.NO", student.reg_num or 'N/A',
            "Mentor name", mentor_name
        ]
        
        for i, text in enumerate(banner_content):
            cell = banner_tbl.cell(0, i)
            cell.text = text
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(255, 255, 255)
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(14)
            p.font.bold = True
            p.font.name = 'Times New Roman'
            p.font.color.rgb = RGBColor(0, 0, 0)
            p.alignment = PP_ALIGN.LEFT if i % 2 == 0 else PP_ALIGN.CENTER

        # Student Photo with gray border
        left_img = Inches(0.5)
        top_img = Inches(2.2)
        width_img = Inches(2.5)
        height_img = Inches(3.0)
        
        frame = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, left_img - Inches(0.1), top_img - Inches(0.1), width_img + Inches(0.2), height_img + Inches(0.2))
        frame.fill.solid()
        frame.fill.fore_color.rgb = RGBColor(230, 230, 230)
        frame.line.color.rgb = RGBColor(200, 200, 200)

        photo_added = False
        try:
            if student.photo_path:
                if student.photo_path.startswith('http'):
                    req = urllib.request.Request(student.photo_path, headers={'User-Agent': 'Mozilla/5.0'})
                    with urllib.request.urlopen(req) as response:
                        image_stream = BytesIO(response.read())
                    slide1.shapes.add_picture(image_stream, left_img, top_img, width=width_img, height=height_img)
                    photo_added = True
                elif os.path.exists(student.photo_path):
                    slide1.shapes.add_picture(student.photo_path, left_img, top_img, width=width_img, height=height_img)
                    photo_added = True
        except Exception as e:
            print(f"Error adding photo to PPT: {e}")

        if not photo_added:
            rect = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, left_img, top_img, width_img, height_img)
            rect.fill.solid()
            rect.fill.fore_color.rgb = RGBColor(100, 100, 100)
            text_frame = rect.text_frame
            text_frame.clear()
            p = text_frame.paragraphs[0]
            p.text = "No Photo"
            p.alignment = PP_ALIGN.CENTER
            p.font.size = Pt(16)
            p.font.bold = True
            p.font.color.rgb = RGBColor(255, 255, 255)

        # --- Dynamic Marks Table: one row per slot ---
        all_slots = slots if slots else ["Slot A"]
        n_data_rows = len(all_slots)
        total_rows  = 1 + n_data_rows
        cols = 4
        table_width  = Inches(9.5)
        table_height = max(Inches(2.0), Inches(0.55 + n_data_rows * 0.55))
        left_tbl = Inches(3.3)
        top_tbl  = Inches(2.2)

        table = slide1.shapes.add_table(total_rows, cols, left_tbl, top_tbl, table_width, table_height).table

        # Header Row
        h_labels = ["Exam", "Total Marks", "Marks Obtained", "Class Average Mark"]
        for i, h in enumerate(h_labels):
            cell = table.cell(0, i)
            cell.text = h
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(112, 173, 71)   # SIMATS Green
            p = cell.text_frame.paragraphs[0]
            p.font.color.rgb = RGBColor(255, 255, 255)
            p.font.bold = True
            p.font.size = Pt(18)
            p.alignment = PP_ALIGN.CENTER

        # Data rows – one row per slot
        for row_idx, slot in enumerate(all_slots, start=1):
            s_marks      = marks_data.get(slot, {})
            test1_val    = str(s_marks.get('test1', '0') or '0')
            avg_val      = str(s_marks.get('avg',   '0') or '0')
            total_m_val  = str(s_marks.get('total_marks', '-') or '-')

            r_vals = [f"Test 1 ({slot})", total_m_val, test1_val, avg_val]
            for c_idx, val in enumerate(r_vals):
                cell = table.cell(row_idx, c_idx)
                cell.text = val
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(226, 239, 218)   # Light green
                p = cell.text_frame.paragraphs[0]
                p.font.size = Pt(18)
                p.font.bold = True
                p.alignment = PP_ALIGN.CENTER

        # --- SLIDE 2: MENTOR NOTES & ATTENDANCE ---
        slide2 = prs.slides.add_slide(slide_layout)

        # Read grade toggle setting
        ppt_settings = GlobalSettings.query.first()
        show_grades = ppt_settings.show_grades_in_ppt if ppt_settings else False

        # Main Gray Content Box
        body_box = slide2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(1.4), Inches(12.3), Inches(5.8))
        body_box.fill.solid()
        body_box.fill.fore_color.rgb = RGBColor(245, 245, 245)
        body_box.line.color.rgb = RGBColor(200, 200, 200)

        tf_body = slide2.shapes.add_textbox(Inches(0.6), Inches(1.5), Inches(12.1), Inches(5.6)).text_frame
        tf_body.word_wrap = True

        # "Welcome to SIMATS ENGINEERING" with green highlight
        p = tf_body.paragraphs[0]
        p.text = "Welcome to SIMATS ENGINEERING"
        p.font.bold = True
        p.font.size = Pt(20)
        p.font.color.rgb = RGBColor(0, 0, 0)
        if len(p.runs) > 0: add_highlight(p.runs[0], '00FF00')

        p = tf_body.add_paragraph()
        p.text = "Dear Parent,"
        p.font.size = Pt(18)
        p.font.bold = True
        p.space_after = Pt(10)

        p = tf_body.add_paragraph()

        low_attendance = False
        for slot in slots:
            try:
                if int(att_data.get(slot, 0)) < 80:
                    low_attendance = True
                    break
            except (ValueError, TypeError):
                pass

        if low_attendance:
            p.text = f"{student.name or 'The student'} has attendance below 80%. Please maintain the attendance % above 80%."
            p.font.color.rgb = RGBColor(0, 0, 0)
            if len(p.runs) > 0: add_highlight(p.runs[0], 'FF0000')
        else:
            p.text = f"So far {student.name or 'the student'} has maintained consistent attendance in the course."
            p.font.color.rgb = RGBColor(0, 0, 0)
            if len(p.runs) > 0: add_highlight(p.runs[0], '00FF00')

        p.font.size = Pt(18)
        p.font.bold = True

        for slot in slots:
            p = tf_body.add_paragraph()
            course_name = marks_data.get(slot, {}).get('course', '')
            if course_name:
                p.text = f"Attendance for {slot}: {course_name}: {att_data.get(slot, 0)}%"
            else:
                p.text = f"Attendance for {slot}: {att_data.get(slot, 0)}%"
            p.font.size = Pt(18)
            p.font.bold = True
            if len(p.runs) > 0: add_highlight(p.runs[0], 'FFFF00')

        # Use student-specific observation if set, otherwise fall back to global observation
        global_obs_record = GlobalMentorObservation.query.first()
        global_obs_text = global_obs_record.content if global_obs_record else 'I personally advised the student to concentrate more on study and skill development.'
        observation_to_use = student.additional_description.strip() if (student.additional_description and student.additional_description.strip()) else global_obs_text
        p = tf_body.add_paragraph()
        p.space_before = Pt(15)
        p.text = observation_to_use
        p.font.size = Pt(18)
        p.font.bold = True
        if len(p.runs) > 0: add_highlight(p.runs[0], 'FFFF00')

        if show_grades:
            # --- GRADES MODE ---
            try:
                results = json.loads(student.results_data) if student.results_data else []
            except:
                results = []
            if results:
                p = tf_body.add_paragraph()
                p.text = "Subject Results:"
                p.font.size = Pt(18)
                p.font.bold = True
                p.space_before = Pt(10)
                if len(p.runs) > 0: add_highlight(p.runs[0], 'FFFF00')
                for r in results:
                    p = tf_body.add_paragraph()
                    p.text = f"  {r.get('subject', '')}  —  Grade: {r.get('grade', '')}"
                    p.font.size = Pt(18)
                    p.font.bold = True
                    if len(p.runs) > 0: add_highlight(p.runs[0], 'FFFF00')
            else:
                p = tf_body.add_paragraph()
                p.text = "No subject results added yet."
                p.font.size = Pt(18)
                p.font.bold = True
                p.space_before = Pt(10)
        else:
            # --- STANDARD MODE ---
            p = tf_body.add_paragraph()
            p.text = f"New course: {student.registered_new_course or 'N/A'}"
            p.font.size = Pt(18)
            p.space_before = Pt(10)
            if len(p.runs) > 0: add_highlight(p.runs[0], 'FFFF00')

            p = tf_body.add_paragraph()
            p.space_before = Pt(15)
            if student.event_participation and student.event_participation.strip():
                p.text = f"Your ward participated in: {student.event_participation.strip()} and gave his very best throughout the journey. His dedication, hard work, and sincere efforts are truly appreciable."
            else:
                p.text = "We encourage your ward to actively participate in upcoming events and extracurricular activities to build their skills and gain valuable experience."
            p.font.size = Pt(18)
            p.font.bold = True
            if len(p.runs) > 0: add_highlight(p.runs[0], 'FFFF00')

        # Add green institutional advisory lines (Global or Custom Override)
        advisory_to_use = student.custom_advisory.strip() if (student.custom_advisory and student.custom_advisory.strip()) else (GlobalAdvisory.query.first().content if GlobalAdvisory.query.first() else "All students are advised to pay their 2nd-year tuition fees on time through the Viana Portal.\nAdditionally, kindly upload your recent passport-size photograph to your Viana profile at the earliest...")

        for adv_line in advisory_to_use.split('\n'):
            if adv_line.strip():
                p = tf_body.add_paragraph()
                p.text = adv_line.strip()
                p.font.size = Pt(18)
                p.font.bold = True
                if len(p.runs) > 0: add_highlight(p.runs[0], '00FF00')

        # Apply Times New Roman font to all paragraphs
        for paragraph in tf_body.paragraphs:
            paragraph.font.name = 'Times New Roman'
            if len(paragraph.runs) > 0:
                paragraph.runs[0].font.name = 'Times New Roman'

    report_path = 'Mentor_Dashboard_Report.pptx'
    prs.save(report_path)
    return send_file(report_path, as_attachment=True)

@app.route('/clear_student/<reg_num>', methods=['POST'])
def clear_student(reg_num):
    if 'user' not in session or session['role'] != 'faculty':
        return redirect(url_for('login'))
    
    student = StudentDetail.query.filter_by(reg_num=reg_num, mentor_username=session['user']).first()
    if student:
        student.slot_info = '[]'
        student.attendance_data = '{}'
        student.marks_data = '{}'
        student.registered_new_course = None
        student.online_course = None
        student.event_participation = None
        student.additional_description = None
        delete_photo(student.photo_path)
        student.photo_path = None
        from datetime import timedelta
        student.last_updated = datetime.utcnow() + timedelta(hours=5, minutes=30)
        db.session.commit()
        flash(f'Report data for {student.name or reg_num} has been cleared.', 'success')
    else:
        flash('Student not found.', 'danger')
    
    return redirect(url_for('faculty_dashboard'))

@app.route('/remove_student/<reg_num>', methods=['POST'])
def remove_student(reg_num):
    if 'user' not in session or session['role'] != 'faculty':
        return redirect(url_for('login'))
    
    student = StudentDetail.query.filter_by(reg_num=reg_num, mentor_username=session['user']).first()
    if student:
        delete_photo(student.photo_path)
        db.session.delete(student)
    
    user = User.query.get(reg_num)
    if user:
        db.session.delete(user)
    
    db.session.commit()
    flash(f'Student {reg_num} has been permanently removed.', 'success')
    return redirect(url_for('faculty_dashboard'))

@app.route('/delete_all_reports', methods=['POST'])
def delete_all_reports():
    if 'user' not in session or session['role'] != 'faculty':
        return redirect(url_for('login'))
    
    students = StudentDetail.query.filter_by(mentor_username=session['user']).all()
    for student in students:
        delete_photo(student.photo_path)
        student.slot_info = '[]'
        student.attendance_data = '{}'
        student.marks_data = '{}'
        student.registered_new_course = None
        student.online_course = None
        student.event_participation = None
        student.additional_description = None
        student.photo_path = None

    db.session.commit()
    flash('All student report data has been cleared. Accounts are still active.', 'success')
    return redirect(url_for('faculty_dashboard'))

if __name__ == '__main__':
    with app.app_context():
        db.create_all()
        seed_db()
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
