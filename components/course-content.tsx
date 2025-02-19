"use client";

import React from "react";
import ReactMarkdown, { Components } from "react-markdown";
import remarkGfm from "remark-gfm";
import { Card } from "@/components/ui/card";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";

const MarkdownComponents: Components = {
  h1: (props) => (
    <h1
      className="scroll-m-20 text-4xl font-extrabold tracking-tight lg:text-5xl mb-4"
      {...props}
    />
  ),
  h2: (props) => (
    <h2
      className="scroll-m-20 border-b pb-2 text-3xl font-semibold tracking-tight first:mt-0 mt-8 mb-4"
      {...props}
    />
  ),
  h3: (props) => (
    <h3
      className="scroll-m-20 text-2xl font-semibold tracking-tight mt-6 mb-3"
      {...props}
    />
  ),
  p: (props) => (
    <p className="leading-7 [&:not(:first-child)]:mt-4" {...props} />
  ),
  ul: (props) => (
    <ul className="my-6 ml-6 list-disc [&>li]:mt-2" {...props} />
  ),
  ol: (props) => (
    <ol className="my-6 ml-6 list-decimal [&>li]:mt-2" {...props} />
  ),
  blockquote: (props) => (
    <blockquote className="mt-6 border-l-2 pl-6 italic" {...props} />
  ),
  table: (props) => (
    <div className="my-6 w-full overflow-y-auto">
      <Table {...props} />
    </div>
  ),
  thead: TableHeader,
  tbody: TableBody,
  tr: TableRow,
  th: (props) => <TableHead {...props} />,
  td: (props) => <TableCell {...props} />,
  pre: (props) => (
    <Card className="my-6 p-4 overflow-x-auto bg-muted">
      <pre className="text-sm" {...props} />
    </Card>
  ),
};

interface CourseContentProps {
  content: string;
}

export function CourseContent({ content }: CourseContentProps) {
  return (
    <div className="prose prose-stone dark:prose-invert max-w-none">
      <ReactMarkdown remarkPlugins={[remarkGfm]} components={MarkdownComponents}>
        {content}
      </ReactMarkdown>
    </div>
  );
}
