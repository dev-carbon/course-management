import {
  Breadcrumb,
  BreadcrumbList,
  BreadcrumbItem,
  BreadcrumbLink,
  BreadcrumbSeparator,
  BreadcrumbPage,
} from "@/components/ui/breadcrumb";
import { Button } from "@/components/ui/button";
import {
  Card,
  CardContent,
  CardDescription,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";
import { getCourses, getDates, getSubjects } from "@/lib/course-utils";
import { format, parse } from "date-fns";
import { Home, Users } from "lucide-react";
import Link from "next/link";
import React from "react";

interface Params {
  subject: string;
  date: string;
}

export function generateStaticParams() {
  const paths: Params[] = [];

  getSubjects().forEach((subject) => {
    getDates(subject).forEach((date) => {
      paths.push({
        subject,
        date,
      });
    });
  });

  return paths;
}

const CoursesPages = async ({ params }: { params: Promise<Params> }) => {
  const { subject, date } = await params;

  const formatedDate = format(
    parse(date, "yyyy-mm-dd", new Date()),
    "yyyy MMMM dd"
  );
  const courses = getCourses(subject, date);

  return (
    <div>
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
              <BreadcrumbPage>{date}</BreadcrumbPage>
            </BreadcrumbItem>
          </BreadcrumbList>
        </Breadcrumb>
      </div>

      <Card>
        <CardHeader>
          <CardTitle>Courses</CardTitle>
          <CardDescription>{formatedDate}</CardDescription>
        </CardHeader>
        <CardContent>
          <Table>
            <TableHeader>
              <TableRow>
                <TableHead>Session</TableHead>
                <TableHead>Titre</TableHead>
                <TableHead>Étudiants</TableHead>
                <TableHead className="text-right">Action</TableHead>
              </TableRow>
            </TableHeader>
            <TableBody>
              {courses.map((course, index) => (
                <TableRow key={index}>
                  <TableCell>Session {course.metadata.sessionNumber}</TableCell>
                  <TableCell>{course.metadata.title}</TableCell>
                  <TableCell>
                    <div className="flex items-center gap-2">
                      <Users className="h-4 w-4" />
                      <span>{course.metadata.students.length}</span>
                    </div>
                  </TableCell>
                  <TableCell className="text-right">
                    <Button variant="outline" asChild>
                      <Link href={`/${subject}/${date}/${course.slug}`}>
                        Open
                      </Link>
                    </Button>
                  </TableCell>
                </TableRow>
              ))}
            </TableBody>
          </Table>
        </CardContent>
      </Card>
    </div>
  );
};

export default CoursesPages;
