CREATE TABLE `jobs` (
	`id` integer PRIMARY KEY AUTOINCREMENT NOT NULL,
	`code` text NOT NULL,
	`quantity` integer DEFAULT 1 NOT NULL,
	`customer` text DEFAULT '' NOT NULL,
	`occasion` text DEFAULT '' NOT NULL,
	`route` text DEFAULT 'undecided' NOT NULL,
	`current_stage` text DEFAULT 'plotter' NOT NULL,
	`status` text DEFAULT 'active' NOT NULL,
	`notes` text DEFAULT '' NOT NULL,
	`created_by` text NOT NULL,
	`created_at` text NOT NULL,
	`updated_at` text NOT NULL,
	`completed_at` text
);
--> statement-breakpoint
CREATE INDEX `jobs_stage_status_idx` ON `jobs` (`current_stage`,`status`);--> statement-breakpoint
CREATE INDEX `jobs_updated_idx` ON `jobs` (`updated_at`);--> statement-breakpoint
CREATE TABLE `stage_events` (
	`id` integer PRIMARY KEY AUTOINCREMENT NOT NULL,
	`job_id` integer NOT NULL,
	`from_stage` text NOT NULL,
	`to_stage` text NOT NULL,
	`actor_email` text NOT NULL,
	`actor_role` text NOT NULL,
	`note` text DEFAULT '' NOT NULL,
	`created_at` text NOT NULL,
	FOREIGN KEY (`job_id`) REFERENCES `jobs`(`id`) ON UPDATE no action ON DELETE cascade
);
--> statement-breakpoint
CREATE INDEX `events_job_idx` ON `stage_events` (`job_id`,`created_at`);--> statement-breakpoint
CREATE TABLE `users` (
	`id` integer PRIMARY KEY AUTOINCREMENT NOT NULL,
	`email` text NOT NULL,
	`display_name` text NOT NULL,
	`role` text DEFAULT 'pending' NOT NULL,
	`created_at` text NOT NULL,
	`updated_at` text NOT NULL
);
--> statement-breakpoint
CREATE UNIQUE INDEX `users_email_uq` ON `users` (`email`);