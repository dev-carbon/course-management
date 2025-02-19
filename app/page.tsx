import { Breadcrumb, BreadcrumbItem, BreadcrumbList, BreadcrumbPage } from "@/components/ui/breadcrumb";
import {
  Card,
  CardContent,
  CardDescription,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";
import { getSubjects } from "@/lib/course-utils";
import { Home } from "lucide-react";
import Link from "next/link";

export default function HomePage() {
  const subjects = getSubjects();

  return (
    <div>
      <div className="mb-6">
        <Breadcrumb>
          <BreadcrumbList>
            <BreadcrumbItem>
              <BreadcrumbPage>
                <Home />
              </BreadcrumbPage>
            </BreadcrumbItem>
          </BreadcrumbList>
        </Breadcrumb>
      </div>

      <Card>
        <CardHeader>
          <CardTitle>Subjects</CardTitle>
          <CardDescription>Select Subject</CardDescription>
        </CardHeader>
        <CardContent>
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
            {subjects.map((subject, index) => (
              <Link key={index} href={`/${subject}`}>
                <Card>
                  <CardHeader>
                    <CardTitle className="text-xl text-center capitalize">
                      {subject}
                    </CardTitle>
                  </CardHeader>
                </Card>
              </Link>
            ))}
          </div>
        </CardContent>
      </Card>
    </div>
  );
}
