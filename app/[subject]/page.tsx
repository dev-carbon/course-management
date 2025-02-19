import {
  Breadcrumb,
  BreadcrumbItem,
  BreadcrumbLink,
  BreadcrumbList,
  BreadcrumbPage,
  BreadcrumbSeparator,
} from "@/components/ui/breadcrumb";
import {
  Card,
  CardContent,
  CardDescription,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";
import { getDates, getSubjects } from "@/lib/course-utils";
import { format, parse } from "date-fns";
import { Home } from "lucide-react";
import Link from "next/link";
import React from "react";

export function generateStaticParams() {
  return getSubjects().map((subject) => ({
    subject,
  }));
}

const SubjectPage = async ({
  params,
}: {
  params: Promise<{ subject: string }>;
}) => {
  const subject = (await params).subject;

  const dates = getDates(subject);

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
              <BreadcrumbPage>{subject}</BreadcrumbPage>
            </BreadcrumbItem>
          </BreadcrumbList>
        </Breadcrumb>
      </div>

      <Card>
        <CardHeader>
          <CardTitle>{subject}</CardTitle>
          <CardDescription>Select Date</CardDescription>
        </CardHeader>
        <CardContent>
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
            {dates.map((date, index) => {
              const formatedDate = format(
                parse(date, "yyyy-mm-dd", new Date()),
                "yyyy MMMM dd"
              );

              return (
                <Link key={index} href={`/${subject}/${date}`}>
                  <Card>
                    <CardHeader>
                      <CardTitle>{formatedDate}</CardTitle>
                    </CardHeader>
                  </Card>
                </Link>
              );
            })}
          </div>
        </CardContent>
      </Card>
    </div>
  );
};

export default SubjectPage;
