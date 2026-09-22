import streamlit as st
import pandas as pd
import plotly.express as px

def show_topic_dashboard():
    # ==================================================
    # CUSTOM CSS
    # ==================================================

    st.markdown("""
    <style>

    .main {
        background-color: #f8fafc;
    }

    .metric-card {
        background: white;
        padding: 20px;
        border-radius: 15px;
        box-shadow: 0px 2px 8px rgba(0,0,0,0.08);
        text-align:center;
    }

    .metric-value{
        font-size:32px;
        font-weight:bold;
        color:#0f766e;
    }

    .metric-label{
        color:#64748b;
        font-size:14px;
    }

    </style>
    """, unsafe_allow_html=True)

    # ==================================================
    # LOAD DATA
    # ==================================================

    FILE_PATH = "Topic_SubTopic_Cancelled_Offline_May2026.xlsx"

    @st.cache_data
    def load_data():

        topic_df = pd.read_excel(
            FILE_PATH,
            sheet_name="Topic-SubTopic"
        )

        cancelled_df = pd.read_excel(
            FILE_PATH,
            sheet_name="Cancelled Sessions"
        )

        offline_df = pd.read_excel(
            FILE_PATH,
            sheet_name="Offline Sessions"
        )

        return topic_df, cancelled_df, offline_df

    topic_df, cancelled_df, offline_df = load_data()

    # ==================================================
    # SIDEBAR
    # ==================================================

    st.sidebar.title("📊 Filters")

    state_filter = st.sidebar.multiselect(
        "State",
        sorted(topic_df["state"].dropna().unique())
    )

    subject_filter = st.sidebar.multiselect(
        "Subject",
        sorted(topic_df["subject"].dropna().unique())
    )

    grade_filter = st.sidebar.multiselect(
        "Grade",
        sorted(topic_df["grade"].dropna().unique())
    )

    if state_filter:
        topic_df = topic_df[
            topic_df["state"].isin(state_filter)
        ]

    if subject_filter:
        topic_df = topic_df[
            topic_df["subject"].isin(subject_filter)
        ]

    if grade_filter:
        topic_df = topic_df[
            topic_df["grade"].isin(grade_filter)
        ]

    # ==================================================
    # TITLE
    # ==================================================



    # ==================================================
    # KPI CARDS
    # ==================================================

    total_sessions = len(topic_df)

    unique_topics = topic_df["topic_name"].nunique()

    unique_subtopics = topic_df["sub_topic_name"].nunique()

    unique_centers = topic_df["center_id"].nunique()

    col1,col2,col3,col4 = st.columns(4)

    with col1:
        st.metric("Total Sessions", f"{total_sessions:,}")

    with col2:
        st.metric("Unique Topics", f"{unique_topics:,}")

    with col3:
        st.metric("Unique Subtopics", f"{unique_subtopics:,}")

    with col4:
        st.metric("Active Centers", f"{unique_centers:,}")

    # ==================================================
    # TABS
    # ==================================================

    tab1, tab2, tab3 = st.tabs([
        "📚 Topic Analysis",
        "🚫 Cancelled Sessions",
        "📵 Offline Sessions"
    ])

    # ==================================================
    # TOPIC ANALYSIS
    # ==================================================

    with tab1:

        st.subheader("Top Topics")

        topic_summary = (
            topic_df
            .groupby("topic_name")
            .size()
            .reset_index(name="Sessions")
            .sort_values(
                "Sessions",
                ascending=False
            )
            .head(20)
        )

        fig = px.bar(
            topic_summary,
            x="Sessions",
            y="topic_name",
            orientation="h",
            title="Top 20 Topics"
        )

        st.plotly_chart(
            fig,
            use_container_width=True
        )

        st.subheader("Top Sub Topics")

        subtopic_summary = (
            topic_df
            .groupby("sub_topic_name")
            .size()
            .reset_index(name="Sessions")
            .sort_values(
                "Sessions",
                ascending=False
            )
            .head(20)
        )

        fig2 = px.bar(
            subtopic_summary,
            x="Sessions",
            y="sub_topic_name",
            orientation="h",
            title="Top 20 Sub Topics"
        )

        st.plotly_chart(
            fig2,
            use_container_width=True
        )

        st.subheader("Subject Distribution")

        subject_chart = (
            topic_df
            .groupby("subject")
            .size()
            .reset_index(name="Count")
        )

        fig3 = px.pie(
            subject_chart,
            names="subject",
            values="Count"
        )

        st.plotly_chart(
            fig3,
            use_container_width=True
        )

    # ==================================================
    # CANCELLED
    # ==================================================

    with tab2:

        st.subheader("Cancelled Session Analysis")

        st.metric(
            "Total Cancelled Sessions",
            len(cancelled_df)
        )

        cancel_chart = (
            cancelled_df
            .groupby("cancel_reason")
            .size()
            .reset_index(name="Count")
            .sort_values(
                "Count",
                ascending=False
            )
            .head(15)
        )

        fig4 = px.bar(
            cancel_chart,
            x="Count",
            y="cancel_reason",
            orientation="h",
            title="Cancellation Reasons"
        )

        st.plotly_chart(
            fig4,
            use_container_width=True
        )

        state_cancel = (
            cancelled_df
            .groupby("State")
            .size()
            .reset_index(name="Count")
            .sort_values(
                "Count",
                ascending=False
            )
        )

        fig5 = px.bar(
            state_cancel,
            x="State",
            y="Count",
            title="Cancelled Sessions by State"
        )

        st.plotly_chart(
            fig5,
            use_container_width=True
        )

    # ==================================================
    # OFFLINE
    # ==================================================

    with tab3:

        st.subheader("Offline Session Analysis")

        st.metric(
            "Total Offline Sessions",
            len(offline_df)
        )

        offline_reason = (
            offline_df
            .groupby("Offline reason")
            .size()
            .reset_index(name="Count")
            .sort_values(
                "Count",
                ascending=False
            )
        )

        fig6 = px.bar(
            offline_reason,
            x="Count",
            y="Offline reason",
            orientation="h",
            title="Offline Reasons"
        )

        st.plotly_chart(
            fig6,
            use_container_width=True
        )

        state_offline = (
            offline_df
            .groupby("State")
            .size()
            .reset_index(name="Count")
            .sort_values(
                "Count",
                ascending=False
            )
        )

        fig7 = px.bar(
            state_offline,
            x="State",
            y="Count",
            title="Offline Sessions by State"
        )

        st.plotly_chart(
            fig7,
            use_container_width=True
        )

    # ==================================================
    # DOWNLOAD
    # ==================================================

    st.download_button(
        "📥 Download Topic Data",
        topic_df.to_csv(index=False),
        file_name="topic_data.csv"
    )
