import fs from "fs";
import path from "path";
import { parse, isValid, format } from "date-fns";
import matter from "gray-matter";

const CONTENT_DIR = "content";

export interface CourseMeta {
  title: string;
  date: string;
  students: string[];
  sessionNumber: number;
  files: string[];
}

export interface Course {
  metadata: CourseMeta,
  content: string,
  slug: string
}

export function getSubjects(): string[] {
  return fs.readdirSync(path.join(process.cwd(), CONTENT_DIR));
}

export function getDates(subject: string): string[] {
  try {
    const dirPath = path.join(process.cwd(), CONTENT_DIR, subject);
    const files = fs.readdirSync(dirPath);

    return files
      .filter((file) => /^\d{4}-\d{2}-\d{2}/.test(file))
      .map((file) => ({
        name: file,
        date: parse(file.substring(0, 10), "yyyy-MM-dd", new Date()),
      }))
      .filter(({ date }) => isValid(date))
      .sort((a, b) => a.date.getTime() - b.date.getTime())
      .map(({ name }) => name);
  } catch (error) {
    console.error("Erreur lors de la lecture du repertoire", error);
    return [];
  }
}

export function getCourses(subject: string, date: string): Course[] {
  try {
    const coursePath = path.join(
      process.cwd(),
      CONTENT_DIR,
      subject,
      date
    );
    const files = fs.readdirSync(coursePath);

    return files
      .filter((file) => file.endsWith(".md"))
      .map((file) => {
        const filePath = path.join(coursePath, file);
        const fileContent = fs.readFileSync(filePath);
        const { data, content } = matter(fileContent);

        return {
          metadata: {
            title: data.title || "Untitled",
            date: format(parse(date, "yyyy-mm-dd", new Date()), "yyyy MMMM dd"),
            students: data.students || [],
            sessionNumber: parseInt(file.split("-")[1].replace(".md", "")),
            files: data.files || []
          },
          content,
          slug: file.replace(".md", ""),
        };
      })
      .sort((a, b) => a.metadata.sessionNumber - b.metadata.sessionNumber);
  } catch (error) {
    console.error("Erreur lors de la lecture des fichiers", error);
    return [];
  }
}

export function getCourseContent(subject: string, date: string, session: string): Course {
  const filePath = path.join(process.cwd(), CONTENT_DIR, subject, date, `${session}.md`);
  const fileContent = fs.readFileSync(filePath, 'utf-8');
  const { data, content } = matter(fileContent);
  
  return {
    metadata: {
      title: data.title || 'Untitled',
      date: format(parse(date, 'yyyy-MM-dd', new Date()), 'dd MMMM yyyy'),
      students: data.students || [],
      sessionNumber: parseInt(session.split('-')[1]),
      files: data.files || []
    },
    content,
    slug: session
  };
}