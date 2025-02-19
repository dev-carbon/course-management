import { CourseContent } from "@/components/course-content";
import { Badge } from "@/components/ui/badge";
import {
  Breadcrumb,
  BreadcrumbItem,
  BreadcrumbLink,
  BreadcrumbList,
  BreadcrumbPage,
  BreadcrumbSeparator,
} from "@/components/ui/breadcrumb";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import {
  getCourseContent,
  getCourses,
  getDates,
  getSubjects,
} from "@/lib/course-utils";
import { Home, Users } from "lucide-react";

interface Params {
  subject: string;
  date: string;
  session: string;
}

export function generateStaticParams() {
  const paths: Params[] = [];

  getSubjects().forEach((subject) => {
    getDates(subject).forEach((date) => {
      getCourses(subject, date).forEach((course) => {
        paths.push({
          subject,
          date,
          session: course.slug,
        });
      });
    });
  });
  return paths;
}

const CoursePage = async ({ params }: { params: Promise<Params> }) => {
  const { subject, date, session } = await params;
  const course = getCourseContent(subject, date, session);

  return (
    <div className="container mx-auto py-8 px-4">
      <div className="mb-6">
        <Breadcrumb>
          <BreadcrumbList>
            <BreadcrumbItem>
              <BreadcrumbLink href="/">
                <Home />
              </BreadcrumbLink>
            </BreadcrumbItem>
            <BreadcrumbSeparator />
            <BreadcrumbItem>
              <BreadcrumbLink href={`/${subject}`}>{subject}</BreadcrumbLink>
            </BreadcrumbItem>
            <BreadcrumbSeparator />
            <BreadcrumbItem>
              <BreadcrumbLink href={`/${subject}/${date}`}>
                {date}
              </BreadcrumbLink>
            </BreadcrumbItem>
            <BreadcrumbSeparator />
            <BreadcrumbItem>
              <BreadcrumbPage>{course.slug}</BreadcrumbPage>
            </BreadcrumbItem>
          </BreadcrumbList>
        </Breadcrumb>
      </div>

      <Card>
        <CardHeader>
          <div className="flex items-center justify-between">
            <CardTitle className="text-3xl">{course.metadata.title}</CardTitle>
            <div className="flex items-center gap-4">
              <div className="flex items-center gap-2">
                <Users className="h-5 w-5" />
                <span className="text-sm text-muted-foreground">
                  {course.metadata.students.length} étudiants
                </span>
              </div>
            </div>
          </div>
          <div className="flex flex-wrap gap-2 mt-4">
            {course.metadata.students.map((student) => (
              <Badge key={student} variant="secondary">
                {student}
              </Badge>
            ))}
          </div>
        </CardHeader>
        <CardContent>
          <CourseContent content={course.content} />
        </CardContent>
      </Card>
    </div>
  );
};

export default CoursePage;
